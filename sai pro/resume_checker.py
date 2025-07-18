import os
import re
import pandas as pd
from docx import Document
import pypdf
import spacy
import warnings
from pyresparser import ResumeParser
from datetime import datetime, timedelta
import time
import traceback
import hashlib
from fuzzywuzzy import fuzz

# --- Outlook specific imports ---
try:
    import win32com.client # Requires pywin32, Windows only
except ImportError:
    print("⚠️ WARNING: 'pywin32' library not found. Outlook integration will not work.")
    print("   Please install it using: pip install pywin32")
    win32com = None

# Suppress specific future warnings from pandas or openpyxl if they occur
warnings.simplefilter(action='ignore', category=FutureWarning)

# NEW: Suppress the specific SpaCy warning about under-constrained version requirement
warnings.filterwarnings("ignore", message=r"\[W094\].*", module="spacy")


# --- Configuration Section: PLEASE EDIT THESE VARIABLES ---
# ==============================================================================
# 1. REPLACE THIS WITH THE ACTUAL PATH TO YOUR FOLDER CONTAINING RESUMES
#    This will now be used as a temporary folder for downloaded resumes.
#    IMPORTANT: Use a raw string (r"...") or forward slashes ("/")
resume_download_folder = r"C:\Users\nanda\Desktop\sai pro\downloaded_resumes" # <--- IMPORTANT: CHANGE THIS PATH!

# 2. Name of your primary output Excel file
excel_file_name = "Resume_Database.xlsx"

# 3. Name of the additional Excel file with candidate details
CADATE_EXCEL_FILE_NAME = "cadate resume details.xlsx" # <--- Name for the second Excel file

# 4. Define the DIRECTORY where you want to save the output Excel files
output_directory = r"C:\Users\nanda\Desktop\sai pro" # <--- IMPORTANT: CHANGE THIS PATH!

# 5. Construct the full path for the primary output Excel file
output_excel_file = os.path.join(output_directory, excel_file_name)

# --- Outlook Specific Configurations ---
OUTLOOK_MAILBOX_NAME = "nanda" # <--- IMPORTANT: Your Outlook mailbox name if different from default "Mailbox - YourName"
INBOX_FOLDER = "Inbox" # <--- Or "Mailbox", "Personal Folders", etc.
RESUME_KEYWORDS_IN_SUBJECT = ["resume", "cv", "application", "job application", "c.v.", "bio-data"] # Keywords to look for in email subject
RESUME_KEYWORDS_IN_BODY = ["resume", "cv", "curriculum vitae", "job application", "attached my resume", "see my attached cv", "application for the position of", "applying for"] # Keywords to look for in email body
RESUME_KEYWORDS_IN_ATTACHMENT_NAME = ["resume", "cv", "application", "profile", "bio", "curriculum_vitae", "cv_"] # Keywords to look for in attachment filenames
RESUME_ATTACHMENT_EXTENSIONS = [".pdf", ".docx", ".doc"] # Allowed resume file extensions
# ==============================================================================


# --- Initialize spaCy ---
try:
    nlp = spacy.load("en_core_web_sm")
except OSError:
    print("\n📦 SpaCy model 'en_core_web_sm' not found. Attempting to download...")
    print("   This might take a moment. Please ensure you have an internet connection.")
    try:
        os.system("python -m spacy download en_core_web_sm")
        nlp = spacy.load("en_core_web_sm")
        print("✅ SpaCy model 'en_core_web_sm' downloaded successfully.")
    except Exception as download_error:
        print(f"❌ ERROR: Failed to download SpaCy model: {download_error}")
        print("   Please try running 'python -m spacy download en_core_web_sm' manually in your terminal.")
        print("   Exiting script. Please resolve SpaCy model issue.")
        exit()


# --- Helper Functions ---
def sanitize_string_for_print(s):
    """Removes or replaces characters that might cause encoding errors during printing."""
    if not isinstance(s, str):
        return str(s) # Convert to string if not already
    return s.encode('utf-8', errors='replace').decode('utf-8')

def extract_text_from_pdf(pdf_path):
    text = ""
    try:
        reader = pypdf.PdfReader(pdf_path)
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            text += page.extract_text()
    except pypdf.errors.PdfReadError as e:
        print(f"  ❌ PDF Read Error: {os.path.basename(pdf_path)} - {e}")
        print("     This might indicate a corrupted or unreadable PDF file.")
    except Exception as e:
        print(f"  ❌ Unexpected Error reading PDF {os.path.basename(pdf_path)}: {e}")
    return text

def extract_text_from_docx(docx_path):
    text = ""
    try:
        doc = Document(docx_path)
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
    except Exception as e:
        print(f"  ❌ Error reading DOCX {os.path.basename(docx_path)}: {e}")
    return text

def convert_doc_to_docx(doc_path):
    """Converts a .doc file to .docx using Microsoft Word (Windows only)."""
    if win32com is None:
        print(f"  ⚠️ Cannot convert .doc to .docx: pywin32 not installed.")
        return None
    
    # Ensure Word is not already open to avoid issues
    word_app = None
    try:
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False # Keep Word invisible
        word_app.DisplayAlerts = False # Suppress alerts

        doc = word_app.Documents.Open(doc_path, ConfirmConversions=False, ReadOnly=True)
        
        # Create a new path for the .docx file in the same directory
        docx_path = os.path.splitext(doc_path)[0] + ".docx"
        
        # wdFormatXMLDocument = 12 (for .docx format)
        doc.SaveAs2(docx_path, FileFormat=12)
        doc.Close()
        print(f"  ✅ Converted '{os.path.basename(doc_path)}' to '{os.path.basename(docx_path)}'")
        return docx_path
    except Exception as e:
        print(f"  ❌ Error converting .doc to .docx for '{os.path.basename(doc_path)}': {e}")
        print("     Ensure Microsoft Word is installed and accessible.")
        return None
    finally:
        if word_app:
            try:
                word_app.Quit()
            except Exception as quit_err:
                print(f"  ⚠️ Warning: Error quitting Word application: {quit_err}")


def is_plausible_name(name_str):
    if not isinstance(name_str, str) or not name_str.strip():
        return False
    name_str = name_str.strip()
    words = name_str.split()
    if len(words) < 1: # A name should have at least one word
        return False
    
    # Check for single-word names (e.g., "Cher") - must be capitalized and not too short
    if len(words) == 1:
        if not (name_str[0].isupper() and len(name_str) > 2 and not name_str.isupper()):
            return False # Reject "cv", "resume", "a", "b", etc.

    if len(words) > 6: # Too many words usually means it's not just a name
        return False
    
    # Reject if it contains numbers or too many special characters
    if bool(re.search(r'\d{3,}', name_str)) or len(re.findall(r'[@#$%^&*()_+={}\[\]|\\:;"<>,?/~`!]', name_str)) > 1:
        return False
    
    # Check for common non-names or resume section headers
    common_headers = ["resume", "curriculum vitae", "contact information", "personal details", "profile", "summary",
                      "experience", "education", "skills", "about me", "references", "work experience", "professional summary",
                      "achievements", "projects", "interests", "cv", "bio", "profile", "contact", "email", "phone", "linkedin"]
    if name_str.lower() in common_headers:
        return False
    
    # Check if all parts are mostly capitalized or common connector words
    if not all(word[0].isupper() or word.lower() in ['de', 'van', 'la', 'del', 'da', 'di', 'du', 'and', 'the'] for word in words if word):
        return False
    
    # Reject if the entire string (cleaned) is all lowercase and multiple words
    if len(name_str.replace(" ", "").replace("-", "").replace(".", "")) > 0 and \
       name_str.replace(" ", "").replace("-", "").replace(".", "").islower() and len(words) > 1:
        return False

    return True

def extract_name_from_filename(filename):
    base_name = os.path.splitext(filename)[0]
    # Remove common resume-related keywords and numbers/copies
    base_name = re.sub(r'\s*\(?\d+\)?$|\s*[-_]?copy\s*$', '', base_name, flags=re.IGNORECASE).strip()
    base_name = re.sub(r'^\W*(resume|cv|bio|profile)\W*|\W*(resume|cv|bio|profile)\W*$', '', base_name, flags=re.IGNORECASE).strip()
    
    # Split by common delimiters
    name_parts = re.split(r'[_\-\s\.]+', base_name)
    
    filtered_parts = []
    for part in name_parts:
        if len(part) > 1 and \
           not part.isdigit() and \
           part.lower() not in ['final', 'new', 'old', 'updated', 'latest', 'version', 'v', 'doc', 'pdf', 'docx', 'for'] and \
           len(part) < 25: # Limit part length to avoid random long strings
            filtered_parts.append(part.capitalize()) # Capitalize for consistency

    extracted_name = " ".join(filtered_parts).strip()
    return extracted_name # is_plausible_name check will be done by caller

def extract_name_from_email(email):
    if not isinstance(email, str) or '@' not in email:
        return None
    local_part = email.split('@')[0]
    # Remove common numerical suffixes (e.g., john.doe123 -> john.doe)
    local_part = re.sub(r'\d+$', '', local_part)
    local_part = re.sub(r'[^\w\.]', '', local_part) # Keep only word chars and dots

    name_parts = re.split(r'[\._\-]+', local_part)
    filtered_parts = [part.capitalize() for part in name_parts if len(part) > 1 and not part.isdigit()] # Filter out single chars/digits
    extracted_name = " ".join(filtered_parts).strip()
    
    return extracted_name # is_plausible_name check will be done by caller

def extract_name_from_email_subject(subject_line):
    name_candidates = []
    
    # Remove common prefixes/suffixes and job titles
    clean_subject = re.sub(r'(?i)\b(resume|cv|application|job application|c\.v\.|bio-data|for the position of|applicant|candidate|re|fw|fwd|fwd:|re:|fw:|fwd)\b', '', subject_line)
    clean_subject = re.sub(r'(?i)\b(software engineer|data scientist|manager|developer|analyst|specialist|intern|associate)\b', '', clean_subject)
    clean_subject = re.sub(r'[\W_]+', ' ', clean_subject).strip() # Replace non-alphanumeric with spaces
    
    # Try to find a name pattern (e.g., two or three capitalized words)
    name_pattern = r'\b([A-Z][a-z]+(?: [A-Z][a-z]+){1,3})\b' # 2-4 words, each capitalized
    matches = re.findall(name_pattern, clean_subject)
    for match in matches:
        name_candidates.append(match)

    # Fallback: Just capitalize and check plausibility of the whole cleaned subject
    potential_name_parts = [word.capitalize() for word in clean_subject.split() if word]
    potential_name = " ".join(potential_name_parts)
    name_candidates.append(potential_name)

    # Use spacy for subject too for PERSON entities if subject is long enough
    if len(subject_line) > 10:
        doc_subject = nlp(subject_line)
        for ent in doc_subject.ents:
            if ent.label_ == "PERSON":
                name_candidates.append(ent.text)

    # Select the longest plausible name found in the subject
    # This function now just returns a strong candidate, final plausibility and selection will be done by caller
    plausible_candidates = [n for n in name_candidates if is_plausible_name(n)]
    if plausible_candidates:
        return max(plausible_candidates, key=len) # Prefer longer names
    return None

def extract_name_from_email_body(body_text):
    name_candidates = []
    lines = body_text.split('\n')
    
    # Check first few lines (salutations, introduction)
    for i, line in enumerate(lines[:min(len(lines), 10)]): # Check first 10 lines max
        line_clean = line.strip()
        if not line_clean: continue

        # Salutations: "Dear [Name]", "Hi [Name]"
        match_salutation = re.search(r'(?i)(Dear|Hi|Hello|Greetings)[,\s:]+\s*([A-Z][a-z]+(?: [A-Z][a-z]+)?(?: [A-Z][a-z]+)?)\b', line_clean)
        if match_salutation:
            name_candidates.append(match_salutation.group(2))
        
        # Introduction: "My name is X Y"
        match_my_name = re.search(r'(?i)(?:My name is|I am)\s+([A-Z][a-z]+(?: [A-Z][a-z]+){1,3})\b', line_clean)
        if match_my_name:
            name_candidates.append(match_my_name.group(1))

        # First line might just be the name
        if i == 0:
            name_candidates.append(line_clean)

    # Check last few lines (signatures)
    for i, line in enumerate(reversed(lines[-min(len(lines), 10):])): # Check last 10 lines max
        line_clean = line.strip()
        if not line_clean: continue
        
        # Signature: "Sincerely, [Name]"
        match_signature = re.search(r'(?i)(Sincerely|Regards|Best regards|Thanks|Yours respectfully)[,\s]*\s*([A-Z][a-z]+(?: [A-Z][a-z]+)?(?: [A-Z][a-z]+)?)\s*$', line_clean)
        if match_signature:
            name_candidates.append(match_signature.group(2))

        # Direct name at the very end
        if i == 0: # Last non-empty line is often just the name
            name_candidates.append(line_clean)

    # Use spacy for body too for PERSON entities (limit text to save processing)
    body_doc = nlp(body_text[:1500]) # Process first 1500 chars for names
    for ent in body_doc.ents:
        if ent.label_ == "PERSON":
            name_candidates.append(ent.text)

    # Select the longest plausible name found in the body
    # This function now just returns a strong candidate, final plausibility and selection will be done by caller
    plausible_candidates = [n for n in name_candidates if is_plausible_name(n)]
    if plausible_candidates:
        return max(plausible_candidates, key=len)
    return None


# UPDATED FUNCTION: Assigns a confidence score to a name based on its source
def get_name_confidence(name, source_type):
    score = 0
    if not is_plausible_name(name):
        return 0 # Not a plausible name at all

    # Base scores - ordered by general reliability
    if source_type == "pyresparser":
        score = 5.0 # Highest confidence, directly from dedicated resume parser
    elif source_type == "basic_parser_resume_text":
        score = 4.5 # High confidence, SpaCy PERSON entity from resume text
    elif source_type == "email_sender_display_name": # NEW: Sender's display name from email
        score = 4.3 # Very high confidence, user-defined name
    elif source_type == "email_body_context": 
        score = 4.0 # High confidence if found in salutations/signatures
    elif source_type == "email_subject_context": 
        score = 3.8 # Good confidence if from clear patterns in subject
    elif source_type == "filename":
        score = 3.0 # Moderate confidence, can be accurate but also generic
    elif source_type == "email_id":
        score = 2.0 # Lowest confidence, prone to nicknames/non-names
    
    # Boost for names with multiple words (suggests a full name)
    if name.count(' ') >= 1:
        score += 0.5 
    
    # Boost for names with reasonable length (not too short, not too long)
    name_len = len(name)
    if 5 <= name_len <= 30: # Typical human name length range
        score += 0.2
    
    return score


def parse_resume_data_basic(text):
    name = "N/A"
    email = "N/A"
    phone = "N/A"
    skills = "N/A"
    experience = "N/A"

    lines = text.split('\n')
    text_for_name_search = "\n".join(lines[:10]) # Limit to top 10 lines for name
    if len(text_for_name_search) > 4000:
        text_for_name_search = text_for_name_search[:4000]

    doc_name = nlp(text_for_name_search)

    all_name_candidates_from_resume = []

    for ent in doc_name.ents:
        if ent.label_ == "PERSON":
            extracted_name = ent.text.strip()
            # is_plausible_name check will be done by the main name selection logic later
            all_name_candidates_from_resume.append(extracted_name)
    
    # Also look for capitalized lines in the top section that could be names
    for line in lines[:5]: # Consider top 5 lines for a direct name line
        line_clean = line.strip()
        if not line_clean:
            continue
        # Remove common job titles/contact info from line before checking
        temp_cleaned_line = re.sub(r'(\s*-\s*|\s*\|\s*|\s*,\s*)\s*(software engineer|data scientist|manager|developer|analyst|specialist|contact|email|phone|profile|summary|experience|education|skills|cv|resume|sr\.|jr\.)\W*', '', line_clean, flags=re.IGNORECASE).strip()
        # is_plausible_name check will be done by the main name selection logic later
        all_name_candidates_from_resume.append(temp_cleaned_line)
    
    # Return all plausible names from basic parser for consideration
    final_name_candidate = "N/A"
    plausible_filtered_names = [n for n in all_name_candidates_from_resume if is_plausible_name(n)]
    if plausible_filtered_names:
        # Prioritize multi-word names, then longest
        multi_word_names = [n for n in plausible_filtered_names if n.count(' ') >= 1]
        if multi_word_names:
            final_name_candidate = max(multi_word_names, key=len)
        else:
            final_name_candidate = max(plausible_filtered_names, key=len)
    
    name = final_name_candidate

    email_match = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', text)
    if email_match:
        email = email_match.group(0).strip().lower()

    phone_patterns = [
        r'(\+?\d{1,4}[-.\s]?)?(\(?\d{2,5}\)?[-.\s]?)?(\d{2,5}[-.\s]?\d{3,4}|\d{7,10})\b',
        r'\b(?:\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})\b',
        r'\b\d{7,15}\b'
    ]
    for pattern in phone_patterns:
        phone_match = re.search(pattern, text)
        if phone_match:
            matched_phone = phone_match.group(0)
            cleaned_phone = re.sub(r'\D', '', matched_phone)
            if matched_phone.startswith('+') and not cleaned_phone.startswith('+'):
                phone = '+' + cleaned_phone
            else:
                phone = cleaned_phone
            if len(phone.replace('+', '')) >= 7:
                if phone.startswith('+91') and len(phone) == 13:
                    phone = phone[3:]
                elif phone.startswith('91') and len(phone) == 12:
                    phone = phone[2:]
                break
            else:
                phone = "N/A"

    current_year = datetime.now().year
    experience_match = re.search(r'(\d+\+?\s*(?:years?|yrs?|yr)\s*(?:of)?\s*(?:(?:overall|total|professional)?\s*(?:experience|exp)))', text, re.IGNORECASE)
    if experience_match:
        experience = experience_match.group(0).strip()
    else:
        experience_text = ""
        experience_section_keywords = [
            r'EXPERIENCE', r'WORK EXPERIENCE', r'PROFESSIONAL EXPERIENCE',
            r'EMPLOYMENT HISTORY', r'JOB HISTORY'
        ]
        
        experience_start_index = -1
        for keyword in experience_section_keywords:
            match = re.search(keyword, text, re.IGNORECASE)
            if match:
                experience_start_index = match.end()
                break

        if experience_start_index != -1:
            sections_after_experience = [
                r'EDUCATION', r'SKILLS', r'PROJECTS', r'AWARDS', r'CERTIFICATIONS',
                r'PUBLICATIONS', r'VOLUNTEER EXPERIENCE', r'REFERENCES', r'INTERESTS'
            ]
            experience_end_index = len(text)

            temp_text_after_exp = text[experience_start_index:]
            for keyword in sections_after_experience:
                match = re.search(keyword, temp_text_after_exp, re.IGNORECASE)
                if match:
                    experience_end_index = experience_start_index + match.start()
                    break
            
            experience_text = text[experience_start_index:experience_end_index]

        if experience_text:
            date_patterns = [
                r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*(?:19|20)\d{2}\s*(?:-|\s+to\s+)\s*(?:(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*(?:19|20)\d{2}|Present|Current|Till Date)',
                r'(?:19|20)\d{2}\s*(?:-|\s+to\s+)\s*(?:(?:19|20)\d{2}|Present|Current|Till Date)'
            ]

            total_months = 0
            
            for pattern in date_patterns:
                for match in re.finditer(pattern, experience_text, re.IGNORECASE):
                    date_range_str = match.group(0)
                    
                    start_month, start_year = None, None
                    end_month, end_year = None, None

                    month_year_pattern = r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s*((?:19|20)\d{2})'
                    year_pattern = r'((?:19|20)\d{2})'

                    dates_found = []

                    for my_match in re.finditer(month_year_pattern, date_range_str, re.IGNORECASE):
                        month_str = my_match.group(1)[:3]
                        year_val = int(my_match.group(2))
                        month_num = {'jan':1, 'feb':2, 'mar':3, 'apr':4, 'may':5, 'jun':6, 'jul':7, 'aug':8, 'sep':9, 'oct':10, 'nov':11, 'dec':12}.get(month_str.lower(), 1)
                        dates_found.append((year_val, month_num))
                    
                    if not dates_found:
                        for y_match in re.finditer(year_pattern, date_range_str, re.IGNORECASE):
                            year_val = int(y_match.group(1))
                            dates_found.append((year_val, 1))

                    if re.search(r'Present|Current|Till Date', date_range_str, re.IGNORECASE):
                        end_year = current_year
                        end_month = datetime.now().month
                        if dates_found:
                            start_year, start_month = dates_found[0]
                    elif len(dates_found) >= 2:
                        start_year, start_month = dates_found[0]
                        end_year, end_month = dates_found[1] # Corrected this line to use dates_found[1]
                    elif len(dates_found) == 1:
                         start_year = dates_found[0][0]
                         start_month = dates_found[0][1]
                         end_year = current_year
                         end_month = datetime.now().month

                    if start_year and end_year:
                        if start_month is None: start_month = 1
                        if end_month is None: end_month = 1

                        duration_months = (end_year - start_year) * 12 + (end_month - start_month) + 1
                        if duration_months > 0:
                            total_months += duration_months

            if total_months > 0:
                total_years = total_months / 12
                if total_years == int(total_years):
                    experience = f"{int(total_years)} years"
                else:
                    experience = f"{total_years:.1f} years"
            
    predefined_skills = [

        # --- VLSI / Semiconductor Skills ---
        # Digital Design (DD) - (Existing, but keeping for context/clarity)
        "Verilog", "VHDL", "SystemVerilog", "RTL Design", "Logic Synthesis",
        "Static Timing Analysis", "STA", "Formal Verification", "Linting",
        "Clock Domain Crossing", "CDC", "Reset Domain Crossing", "RDC",
        "Low Power Design", "Power Analysis", "FPGA Design", "ASIC Design",
        "Combinational Logic", "Sequential Logic", "Finite State Machines", "FSM",
        "Pipelining", "Data Paths", "Control Paths", "Memory Design", "SRAM", "DRAM",
        "I/O Interfaces", "SPI", "I2C", "UART", "ARM Architecture", "RISC-V",
        "Cache Coherence", "High-Level Synthesis", "Synthesis", "Netlist",
        "Timing Closure", "Digital Logic", "Verilog-HDL", "VHDL-AMS",

        # Design Verification (DV) - (Existing, but keeping for context/clarity)
        "UVM", "Universal Verification Methodology", "Specman E", "PSL",
        "SVA", "SystemVerilog Assertions", "Functional Verification",
        "Testbench Architecture", "Test Plan", "Coverage Driven Verification", "CDV",
        "Constrained Random Verification", "CRV", "Model Checking",
        "Assertion-Based Verification", "ABV", "Emulation", "FPGA Prototyping",
        "Regression Management", "Bug Tracking", "Gate Level Simulation", "GLS",
        "Transaction-Level Modeling", "TLM", "Scoreboarding", "Monitors", "Drivers",
        "Sequencers", "Checkers", "Functional Coverage", "Code Coverage",
        "Protocol Verification", "VIP", "Verification IP", "Debugging", "Verdi",
        "VCS", "QuestaSim", "Incisive", "Xcelium", "FormalPro", "SpyGlass",
        "JasperGold", "Symphony", "Unified Power Format", "UPF", "CPF",
        "Verification Methodology", "Verification Plan", "Coverage Closure",
        "Formal Equivalence Checking", "LEC", "Assertions", "Test Automation",

        # Design for Testability (DFT) - (Existing, but keeping for context/clarity)
        "Scan Insertion", "ATPG", "Automatic Test Pattern Generation", "JTAG",
        "Boundary Scan", "MBIST", "Memory Built-In Self-Test", "LBIST",
        "Logic Built-In Self-Test", "Fault Simulation", "Stuck-at Faults",
        "Transition Faults", "Bridging Faults", "Test Compression", "DFT Sign-off",
        "Diagnosis", "ATE", "Automatic Test Equipment", "Scan Chains", "Test Modes",
        "Fault Models", "Test Coverage", "IP-level DFT", "System-level DFT",
        "Delay Testing", "At-speed Testing", "TetraMax", "TestKompress", "DFT Compiler",
        "SMS", "Tessent", "OpTest", "DFTMAX", "DesignWare", "Pattern Generation",
        "Fault Coverage", "Manufacturing Test", "Yield Improvement",

        # Physical Design
        "Physical Design", "Layout Design", "Floorplanning", "Power Grid Network", "PGN",
        "Placement", "Clock Tree Synthesis", "CTS", "Routing", "ECO", "Engineering Change Order",
        "Design Rule Check", "DRC", "Layout Versus Schematic", "LVS", "Parasitic Extraction", "PEX",
        "Power Integrity", "Signal Integrity", "IR Drop Analysis", "EM", "Electromigration",
        "Physical Verification", "DFM", "Design for Manufacturability", "Timing Closure",
        "Cadence Innovus", "Synopsys ICC", "Synopsys ICC2", "Siemens Aprisa", "PrimeTime",
        "Quantus", "Voltus", "Tempus", "Calibre", "StarRC", "NanoRoute", "RedHawk",
        "Chip Assembly", "Tapeout", "GDSII", "LEF", "DEF", "Liberty Format", ".lib",
        "Low Power Implementation", "FinFET", "Process Technology", "Layout Editor",

        # Analog Design
        "Analog Design", "Analog IC Design", "Transistor Level Design", "Schematic Design",
        "Layout Design", "SPICE Simulation", "Noise Analysis", "PVT", "Process Voltage Temperature",
        "Matching", "Bandgap References", "LDO", "Low Dropout Regulator", "PLL", "Phase-Locked Loop",
        "ADC", "Analog-to-Digital Converter", "DAC", "Digital-to-Analog Converter", "Op-Amp",
        "Filters", "Oscillators", "RF Design", "Radio Frequency", "Mixed-Signal Simulation",
        "Cadence Virtuoso", "Spectre", "HSPICE", "Eldo", "ADE", "Analog Design Environment",
        "Mentor Graphics AFS", "Analog FastSPICE", "Keysight ADS", "EMX", "Momentum",
        "Custom Layout", "Device Physics", "CMOS", "Bipolar", "BiCMOS", "Power Management IC",
        "PMIC", "Data Converters", "Amplifiers", "Transceivers", "Analog Front End", "AFE",

        # Analog Mixed-Signal (AMS) Design
        "Analog Mixed-Signal", "AMS Design", "Mixed-Signal Verification", "Co-simulation",
        "Verilog-AMS", "AMS Designer", "Custom Compiler", "Xcelium AMS", "Questa AMS",
        "Behavioral Modeling", "Top-level Integration", "System-level Verification",
        "Spice/FastSpice/UltraSim", "Real-Number Modeling", "RNM", "Mixed Signal Flow",

        # FPGA Design (Enhanced)
        "FPGA", "FPGA Development", "Xilinx Vivado", "AMD Vivado", "Intel Quartus Prime",
        "Altera Quartus", "Lattice Diamond", "Libero SoC", "FPGA Prototyping",
        "Logic Optimization", "IP Integration", "On-chip Debugging", "ILA", "VIO",
        "HLS", "High-Level Synthesis", "Board Bring-up", "System Integration",
        "Synthesis Constraints", "Place and Route", "Timing Closure (FPGA)",
        "FPGA Architecture", "Hardware Description Language", "HDL", "MicroBlaze", "Zynq",
        "NIOS", "Platform Design", "Embedded Processor",

        # IP Design and Characterization
        "IP Design", "IP Core Development", "IP Integration", "IP Verification",
        "IP Hardening", "IP Delivery", "Library Characterization", "Standard Cell Libraries",
        "IO Libraries", "Memory Compilers", "Characterization Tools", "Cadence Liberate",
        "Synopsys SiliconSmart", "Synopsys SiliconSmart", "Liberty (.lib)", "Timing Models",
        "Power Models", "Noise Models", "IP Reuse", "Design IP", "Verification IP", "Test IP",
        "EDA Tools", "Foundry Process", "PDK", "Process Design Kit",

        # Test Chip Development
        "Test Chip", "Test Chip Development", "Silicon Validation", "Post-Silicon Validation",
        "Bring-up", "Debugging", "Characterization", "Measurement", "Yield Analysis",
        "Failure Analysis", "FA", "Wafer Test", "Package Test", "Production Test",
        "ATE Test Program", "Automated Test Equipment", "Parametric Test", "Functional Test",
        "Silicon Debug", "Data Analysis", "Statistical Process Control", "SPC",
        "Product Engineering", "Reliability Testing"
    ]

    found_skills = []
    text_lower = text.lower()
    for skill in predefined_skills:
        # Use word boundaries to avoid partial matches (e.g., "C" matching "C#")
        if re.search(r'\b' + re.escape(skill.lower()) + r'\b', text_lower):
            found_skills.append(skill)

    if found_skills:
        skills = ", ".join(sorted(list(set(found_skills))))

    return {
        "Name": name, # This will be used as a source for 'Candidate Name'
        "Skills": skills, # This will be renamed to 'Skill' later
        "Experience": experience,
        "Email ID": email,
        "Phone Number": phone
    }

# UPDATED Outlook Integration Function
def download_resumes_from_outlook(download_folder, mailbox_name, inbox_name, subject_keywords, body_keywords, attachment_name_keywords, attachment_extensions):
    """
    Connects to Outlook, checks for new emails with resume attachments,
    downloads them, and leaves the emails in the Inbox.
    Returns a list of dictionaries with file_path, received_time, email_subject, email_body, AND email_sender_display_name.
    """
    if win32com is None:
        print("Outlook integration is disabled because 'pywin32' library is not installed.")
        return []

    downloaded_files_info = [] # List to store dictionaries of downloaded file info
    
    os.makedirs(download_folder, exist_ok=True)

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        
        try:
            inbox = outlook.GetDefaultFolder(6) # 6 corresponds to olFolderInbox
            print(f"  📥 Connected to Outlook Inbox: {inbox.FolderPath}")
            time.sleep(1) # Small pause
        except Exception as e:
            print(f"  ❌ Error accessing Outlook Inbox: {e}")
            print("     Attempting to access default Inbox...")
            try:
                inbox = outlook.GetDefaultFolder(6) # olFolderInbox constant
                print(f"  ✅ Successfully accessed default Inbox: {inbox.FolderPath}")
            except Exception as e_default:
                print(f"  ❌ Failed to access default Inbox: {e_default}")
                print("     Outlook Inbox is not accessible. Please ensure Outlook is running and configured correctly.")
                return []

        messages = inbox.Items
        messages.Sort("[ReceivedTime]", False) # Sort by received time, newest first

        # Filter for emails received within the last 1 day (24 hours) - this can be adjusted if needed
        yesterday = datetime.now() - timedelta(days=1) # <--- ADJUST THIS IF YOU NEED TO LOOK FURTHER BACK
        filter_date_str = yesterday.strftime('%m/%d/%Y %H:%M %p') # Format for Outlook filter
        filter_string = f"[ReceivedTime] >= '{filter_date_str}'"
        
        print(f"  ⏳ Filtering emails received after: {filter_date_str}")
        try:
            messages = messages.Restrict(filter_string)
            print(f"  ✅ Filter applied. Checking {messages.Count} email(s).")
        except Exception as filter_error:
            print(f"  ⚠️ WARNING: Could not apply date filter to Outlook messages: {filter_error}")
            print("     Proceeding without date filter (will check all emails in Inbox, which might be slow).")

        email_checked_count = 0
        for message in list(messages): # Convert to list to avoid issues if messages collection changes during loop
            email_checked_count += 1
            
            current_subject = "N/A - Unknown"
            current_sender = "N/A - Unknown" # This will capture the display name
            current_body_snippet = "N/A - Empty" # Store a snippet of the body
            message_received_time = None
            
            try:
                message_received_time = message.ReceivedTime # This is a pywintypes.datetime object
                current_subject = message.Subject if message.Subject else "No Subject"
                current_sender = message.SenderName if message.SenderName else "Unknown Sender" # Capture SenderName
                # Store up to the first 2000 characters of the body for name extraction, prevent memory issues
                current_body_snippet = message.Body[:2000] if message.Body else "" 
            except Exception as read_err:
                sanitized_err = sanitize_string_for_print(str(read_err))
                print(f"  ⚠️ WARNING: Could not read full details for an email (possibly encoding or access issue). Skipping this email. Error: {sanitized_err}")
                continue # Skip to next email if basic info can't be read

            sanitized_subject = sanitize_string_for_print(current_subject)
            sanitized_sender = sanitize_string_for_print(current_sender)
            subject_lower = current_subject.lower()
            body_lower = current_body_snippet.lower() # Use snippet for keyword check too
            
            # Flag to track if any resume attachment was found and downloaded for this specific email
            resume_downloaded_from_this_email = False

            if message.Attachments.Count > 0:
                print(f"  📧 Processing email from '{sanitized_sender}' (Subject: '{sanitized_subject}') - {message.Attachments.Count} attachment(s).")
                for attachment in message.Attachments:
                    attachment_name_safe = "" 
                    original_ext = "" 

                    try:
                        attachment_name_safe = attachment.FileName
                        original_ext = os.path.splitext(attachment_name_safe)[1].lower()
                    except Exception as fn_err:
                        print(f"    ⚠️ Warning: Could not read attachment filename for email '{sanitized_subject}'. Using generic name. Error: {fn_err}")
                        cleaned_subject_for_name = re.sub(r'[^\w\s.-]', '', current_subject).strip()
                        if len(attachment_extensions) > 0:
                             original_ext = attachment_extensions[0] 
                        else:
                            original_ext = ".bin" 
                        attachment_name_safe = f"attachment_from_{cleaned_subject_for_name or 'unknown_subject'}__{int(time.time())}{original_ext}"
                        

                    sanitized_attachment_name_for_print = sanitize_string_for_print(attachment_name_safe)

                    # Check if the attachment itself has a supported extension AND
                    # if the attachment name, OR the email subject, OR the email body contains a resume keyword.
                    if original_ext in attachment_extensions:
                        is_attachment_name_relevant = any(keyword in attachment_name_safe.lower() for keyword in attachment_name_keywords)
                        is_email_content_relevant = any(keyword in subject_lower for keyword in subject_keywords) or \
                                                  any(keyword in body_lower for keyword in body_keywords)

                        if is_attachment_name_relevant or is_email_content_relevant:
                            try:
                                # Construct the save path, handling potential filename duplicates in the download folder
                                save_path = os.path.join(download_folder, attachment_name_safe)
                                base_name_no_ext, current_file_ext = os.path.splitext(save_path)
                                
                                counter = 1
                                while os.path.exists(save_path):
                                    save_path = f"{base_name_no_ext}_{counter}{current_file_ext}"
                                    counter += 1

                                attachment.SaveAsFile(save_path)
                                downloaded_files_info.append({
                                    'file_path': save_path, 
                                    'received_time': message_received_time,
                                    'email_subject': current_subject,
                                    'email_body': current_body_snippet,
                                    'email_sender_display_name': current_sender # Store sender display name
                                })
                                print(f"    📥 Downloaded relevant attachment: {os.path.basename(save_path)}")
                                resume_downloaded_from_this_email = True 
                                
                            except Exception as att_save_err:
                                print(f"    ❌ ERROR: Failed to save attachment '{sanitized_attachment_name_for_print}': {att_save_err}")
                        else:
                            print(f"    ℹ️ Skipping attachment '{sanitized_attachment_name_for_print}' (no strong resume keywords found).")
                    else:
                        print(f"    ℹ️ Skipping attachment '{sanitized_attachment_name_for_print}' (unsupported extension: '{original_ext}').")
            else:
                if any(keyword in subject_lower for keyword in subject_keywords) or \
                   any(keyword in body_lower for keyword in body_keywords):
                    print(f"  ℹ️ Email '{sanitized_subject}' is relevant by text, but has no attachments. Skipping.")

    except Exception as e:
        sanitized_error_overall = sanitize_string_for_print(str(e))
        print(f"❌ CRITICAL ERROR during Outlook processing (overall loop): {sanitized_error_overall}")
        traceback.print_exc() 
    
    return downloaded_files_info 


# --- Main Processing Logic ---
def process_resumes_in_folder(folder_path, excel_file_path, downloaded_files_info):
    """
    Processes resumes in a given folder, parses data, and updates the Excel sheet.
    Implements the new duplicate logic:
    1. If Email OR Phone matches existing, mark as 'Duplicate'.
    2. If Filename AND Skills match existing, DO NOT add.
    """
    print(f"\n📄 Starting resume parsing from: {folder_path}")
    print(f"   Primary output will be saved to: {excel_file_path}")
    print(f"   Secondary output will be saved to: {os.path.join(output_directory, CADATE_EXCEL_FILE_NAME)}")

    resume_data_to_add = [] # List to store dictionaries of resumes that will be added to the DataFrame

    existing_df = pd.DataFrame()
    existing_phone_emails = set() # For checking email/phone duplicates
    existing_filename_skills = set() # For checking filename/skills duplicates for *exclusion*

    if os.path.exists(excel_file_path):
        try:
            existing_df = pd.read_excel(excel_file_path)
            print(f"  Loaded {len(existing_df)} existing records from primary Excel.")

            # Ensure all relevant columns for merge key exist, fill with empty string if not
            for col in ['Candidate Name', 'Phone Number', 'Email ID', 'Skill', 'Total Experience', 'File Name', 'Source Date', 'Status']: 
                # Handle existing 'Name' column if it was present from prior runs, rename to 'Candidate Name'
                if col == 'Candidate Name' and 'Name' in existing_df.columns and 'Candidate Name' not in existing_df.columns:
                    existing_df['Candidate Name'] = existing_df['Name']
                    existing_df.drop(columns=['Name'], inplace=True)
                # Handle existing 'Skills' column if it was present from prior runs, rename to 'Skill'
                elif col == 'Skill' and 'Skills' in existing_df.columns and 'Skill' not in existing_df.columns:
                    existing_df['Skill'] = existing_df['Skills']
                    existing_df.drop(columns=['Skills'], inplace=True)
                # Handle existing 'Received On' column if it was present, rename to 'Source Date'
                elif col == 'Source Date' and 'Received On' in existing_df.columns and 'Source Date' not in existing_df.columns:
                    existing_df['Source Date'] = existing_df['Received On']
                    existing_df.drop(columns=['Received On'], inplace=True)
                # Handle existing 'Experience' column if it was present, rename to 'Total Experience'
                elif col == 'Total Experience' and 'Experience' in existing_df.columns and 'Total Experience' not in existing_df.columns:
                    existing_df['Total Experience'] = existing_df['Experience']
                    existing_df.drop(columns=['Experience'], inplace=True)
                else:
                    # For other columns, ensure they exist with default empty string if missing
                    if col not in existing_df.columns:
                        existing_df[col] = ''
            
            # If 'Date' column exists from previous runs, remove it
            if 'Date' in existing_df.columns:
                existing_df.drop(columns=['Date'], inplace=True)

            # Populate sets for duplicate checks
            for index, row in existing_df.iterrows():
                # For Phone/Email matching
                phone_clean = re.sub(r'\D', '', str(row.get('Phone Number', '')).strip())
                email_clean = str(row.get('Email ID', '')).strip().lower()
                
                if phone_clean and phone_clean != 'n/a':
                    existing_phone_emails.add(phone_clean)
                if email_clean and email_clean != 'n/a':
                    existing_phone_emails.add(email_clean)

                # For Filename/Skills matching (to exclude)
                filename_clean = str(row.get('File Name', '')).strip().lower()
                skills_clean = frozenset(s.strip().lower() for s in str(row.get('Skill', '')).split(',') if s.strip()) 
                
                if filename_clean and skills_clean:
                    existing_filename_skills.add((filename_clean, skills_clean))

        except Exception as e:
            print(f"  ❌ ERROR: Could not load existing Excel file '{excel_file_path}': {e}")
            print("     Starting with an empty sheet.")
            existing_df = pd.DataFrame()
            existing_phone_emails = set()
            existing_filename_skills = set()

    # Create a map from original_file_name to its associated email data
    email_data_map = {os.path.basename(item['file_path']): item for item in downloaded_files_info}

    processed_count = 0
    file_list = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
    if not file_list:
        print("  ℹ️ No files found in the download folder to process.")
        return

    print(f"  Processing {len(file_list)} files...")
    for filename in file_list:
        file_path = os.path.join(folder_path, filename)
        file_extension = os.path.splitext(filename)[1].lower()
        
        original_file_name_for_excel = os.path.basename(file_path) 
        original_file_path_for_cleanup = file_path 

        # Retrieve email subject and body for this file
        file_email_data = email_data_map.get(original_file_name_for_excel, {})
        email_subject = file_email_data.get('email_subject', "N/A")
        email_body = file_email_data.get('email_body', "N/A")
        message_received_time = file_email_data.get('received_time') # Get original received time
        sender_display_name = file_email_data.get('email_sender_display_name', "N/A") # Get sender display name

        # New logic to handle .doc files
        if file_extension == '.doc':
            if win32com:
                converted_docx_path = convert_doc_to_docx(file_path)
                if converted_docx_path:
                    file_path = converted_docx_path 
                    file_extension = '.docx' 
                else:
                    print(f"  ❌ Failed to convert .doc file '{filename}'. Skipping.")
                    continue
            else:
                print(f"  ⏩ Skipping .doc file '{filename}' as pywin32 (and thus MS Word conversion) is not available.")
                continue

        # Now check if the (possibly converted) file's extension is supported
        if file_extension not in [".pdf", ".docx"]:
            print(f"  ⏩ Skipping unsupported file type: {filename}")
            if file_path != original_file_path_for_cleanup and os.path.exists(file_path):
                try: os.remove(file_path)
                except Exception as clean_err: print(f"    ❌ Error cleaning up temp file {os.path.basename(file_path)}: {clean_err}")
            continue

        print(f"\n  --- Processing: {filename} ---")

        extracted_text = ""
        pyresparser_data = {}
        basic_parser_data = {}
        final_parsed_data = {}

        name_from_original_filename = extract_name_from_filename(original_file_name_for_excel)

        try:
            parser = ResumeParser(file_path)
            pyresparser_data = parser.get_extracted_data()

            if pyresparser_data and pyresparser_data.get('mobile_number'):
                 pyresparser_data["mobile_number"] = re.sub(r'\D', '', str(pyresparser_data["mobile_number"])).strip()
            
            if pyresparser_data and pyresparser_data.get('email'):
                pyresparser_data['email'] = str(pyresparser_data['email']).lower().strip()

        except Exception as e:
            print(f"  ⚠️ WARNING: Pyresparser failed for {filename}: {e}. Falling back to basic parsing.")
            pyresparser_data = {}
        
        if file_extension == '.pdf':
            extracted_text = extract_text_from_pdf(file_path)
        elif file_extension == '.docx':
            extracted_text = extract_text_from_docx(file_path)

        if not extracted_text:
            print(f"  ❌ No text extracted from {filename}. Skipping detailed parsing.")
            if file_path != original_file_path_for_cleanup and os.path.exists(file_path):
                try:
                    os.remove(file_path)
                    print(f"    🧹 Cleaned up temporary converted file: {os.path.basename(file_path)}")
                except Exception as clean_err:
                    print(f"    ❌ Error cleaning up temporary file {os.path.basename(file_path)}: {clean_err}")
            continue

        basic_parser_data = parse_resume_data_basic(extracted_text)

        # --- Populate final_parsed_data with best available info ---
        final_parsed_data = {
            "Candidate Name": "N/A", 
            "Skill": "N/A", # Renamed from 'Skills' in basic_parser_data
            "Total Experience": "N/A", # Renamed from 'Experience'
            "Email ID": "N/A", 
            "Phone Number": "N/A", 
            "Source Date": "N/A", 
            "Month": "N/A", 
            "Year": "N/A", 
            "File Name": original_file_name_for_excel
        }

        if message_received_time:
            if not isinstance(message_received_time, datetime):
                try: message_received_time = datetime(message_received_time.year, message_received_time.month, message_received_time.day, message_received_time.hour, message_received_time.minute, message_received_time.second)
                except: message_received_time = None
            if message_received_time:
                final_parsed_data["Source Date"] = message_received_time.strftime('%Y-%m-%d %H:%M:%S')
                final_parsed_data["Month"] = message_received_time.strftime('%B')
                final_parsed_data["Year"] = message_received_time.year

        # --- Name Extraction Logic (Revised with Confidence Scoring and Fuzzy Matching) ---
        name_candidates_with_scores = [] # List of (name, score, source) tuples

        # 1. From resume parsers (highest confidence)
        pyres_name = pyresparser_data.get('name')
        if pyres_name: 
            name_candidates_with_scores.append((pyres_name, get_name_confidence(pyres_name, "pyresparser"), "pyresparser"))

        basic_name = basic_parser_data.get('Name') 
        if basic_name:
            name_candidates_with_scores.append((basic_name, get_name_confidence(basic_name, "basic_parser_resume_text"), "basic_parser_resume_text"))

        # 2. From email sender display name (HIGH CONFIDENCE SOURCE)
        if sender_display_name and sender_display_name != "N/A":
            name_candidates_with_scores.append((sender_display_name, get_name_confidence(sender_display_name, "email_sender_display_name"), "email_sender_display_name"))

        # 3. From filename
        if name_from_original_filename:
            name_candidates_with_scores.append((name_from_original_filename, get_name_confidence(name_from_original_filename, "filename"), "filename"))

        # 4. From email subject
        name_from_subject = extract_name_from_email_subject(email_subject)
        if name_from_subject:
            name_candidates_with_scores.append((name_from_subject, get_name_confidence(name_from_subject, "email_subject_context"), "email_subject_context"))

        # 5. From email body
        name_from_body = extract_name_from_email_body(email_body)
        if name_from_body:
            name_candidates_with_scores.append((name_from_body, get_name_confidence(name_from_body, "email_body_context"), "email_body_context"))

        # 6. From email ID (lowest confidence)
        email_id_candidate = pyresparser_data.get('email') or basic_parser_data.get('Email ID')
        if email_id_candidate and email_id_candidate != "N/A":
            final_parsed_data["Email ID"] = str(email_id_candidate).lower().strip()
            name_from_email_id = extract_name_from_email(final_parsed_data["Email ID"])
            if name_from_email_id:
                name_candidates_with_scores.append((name_from_email_id, get_name_confidence(name_from_email_id, "email_id"), "email_id"))
        
        # Filter out non-plausible names (score 0 indicates not plausible)
        valid_candidates = [(name, score, source) for name, score, source in name_candidates_with_scores if score > 0]
        
        best_name = "N/A"
        if valid_candidates:
            # Sort candidates by confidence score (descending), then by length (descending)
            valid_candidates.sort(key=lambda x: (x[1], len(x[0])), reverse=True)
            
            # Select the best name, avoiding fuzzy duplicates
            selected_best_names_by_group = []
            
            # Keep track of names already "covered" by a selected candidate
            covered_names = set() 

            for current_name, current_score, current_source in valid_candidates:
                # If this name is already very similar to one we've already considered and chosen (higher confidence), skip
                is_already_covered = False
                for existing_covered_name in covered_names:
                    if fuzz.token_sort_ratio(current_name, existing_covered_name) > 85: # High similarity threshold
                        is_already_covered = True
                        break
                
                if not is_already_covered:
                    selected_best_names_by_group.append(current_name)
                    # For simplicity, just add the current name as representative of its group
                    covered_names.add(current_name) 

            # After this loop, selected_best_names_by_group will contain the top candidate
            # from each fuzzy-similar group, ordered by confidence.
            if selected_best_names_by_group:
                # The first one is the overall highest confidence, longest, non-fuzzy-duplicate
                best_name = selected_best_names_by_group[0] 

        final_parsed_data["Candidate Name"] = best_name
        # --- End Name Extraction Logic ---


        pyres_phone_clean = pyresparser_data.get('mobile_number', '')
        basic_phone_clean = basic_parser_data.get('Phone Number', '')

        if pyres_phone_clean and len(pyres_phone_clean) >= 7:
            final_parsed_data["Phone Number"] = pyres_phone_clean
        elif basic_phone_clean and len(basic_phone_clean) >= 7:
            final_parsed_data["Phone Number"] = basic_phone_clean
        
        if final_parsed_data["Phone Number"] != 'N/A' and final_parsed_data["Phone Number"] is not None:
            final_parsed_data["Phone Number"] = re.sub(r'\D', '', str(final_parsed_data["Phone Number"]))
            if not final_parsed_data["Phone Number"]:
                final_parsed_data["Phone Number"] = "N/A"

        pyres_exp = pyresparser_data.get('total_experience')
        basic_exp = basic_parser_data.get('Experience') # Note: basic_parser_data still uses 'Experience'

        if isinstance(pyres_exp, (int, float)) and pyres_exp > 0:
            final_parsed_data["Total Experience"] = f"{int(pyres_exp)} years"
        elif basic_exp != "N/A":
            final_parsed_data["Total Experience"] = basic_exp
        else:
            final_parsed_data["Total Experience"] = "N/A"

        combined_skills_set = set()
        pyresparser_skills_list = [s.strip() for s in pyresparser_data.get('skills', []) if s.strip()]
        if pyresparser_skills_list:
            combined_skills_set.update(pyresparser_skills_list)
        
        basic_skills_list = []
        if isinstance(basic_parser_data.get('Skills'), str) and basic_parser_data.get('Skills') != "N/A":
            basic_skills_list = [s.strip() for s in basic_parser_data['Skills'].split(',') if s.strip()]
        if basic_skills_list:
            combined_skills_set.update(basic_skills_list)

        final_parsed_data["Skill"] = ", ".join(sorted(list(set(combined_skills_set)))) if combined_skills_set else "N/A"
        
        # --- Apply new duplicate logic ---
        # Rule 2: If Filename AND Skills match existing, DO NOT ADD
        current_filename_clean = final_parsed_data.get('File Name', '').strip().lower()
        current_skills_frozenset = frozenset(s.strip().lower() for s in final_parsed_data.get('Skill', '').split(',') if s.strip())

        if (current_filename_clean, current_skills_frozenset) in existing_filename_skills and current_filename_clean and current_skills_frozenset:
            print(f"  🛑 Skipping: '{filename}' - Duplicate Filename AND Skill found in existing data. Not adding.")
            if file_path != original_file_path_for_cleanup and os.path.exists(file_path):
                try: os.remove(file_path)
                except Exception as clean_err: print(f"    ❌ Error cleaning up temp file {os.path.basename(file_path)}: {clean_err}")
            continue 

        # Rule 1: If Email OR Phone matches existing, mark as 'Duplicate'
        current_phone_clean = re.sub(r'\D', '', str(final_parsed_data.get('Phone Number', ''))).strip()
        current_email_clean = str(final_parsed_data.get('Email ID', '')).strip().lower()
        
        is_contact_duplicate = False
        if current_phone_clean and current_phone_clean != 'n/a' and current_phone_clean in existing_phone_emails:
            is_contact_duplicate = True
        if current_email_clean and current_email_clean != 'n/a' and current_email_clean in existing_phone_emails:
            is_contact_duplicate = True

        if is_contact_duplicate:
            final_parsed_data['Status'] = 'Duplicate'
            print(f"  ⏩ Marked as Duplicate: '{filename}' (Email or Phone matches existing record).")
        else:
            final_parsed_data['Status'] = 'New'
            print(f"  ⭐ Marked as New: '{filename}'.")
            
        resume_data_to_add.append(final_parsed_data)
        processed_count += 1
        
        # Clean up the converted .docx file if it was created for a successfully processed file
        if file_path != original_file_path_for_cleanup and os.path.exists(file_path):
            try:
                os.remove(file_path)
                print(f"    🧹 Cleaned up temporary converted file: {os.path.basename(file_path)}")
            except Exception as clean_err:
                print(f"    ❌ Error cleaning up temporary file {os.path.basename(file_path)}: {clean_err}")

    
    if processed_count == 0:
        print("  ℹ️ No new unique resumes processed or added in this run.")
        return

    if resume_data_to_add:
        new_resumes_df = pd.DataFrame(resume_data_to_add)

        if 'Status' not in existing_df.columns:
            existing_df['Status'] = 'Existing' 

        combined_df = pd.concat([existing_df, new_resumes_df], ignore_index=True)

        combined_df['Phone Number_Clean'] = combined_df['Phone Number'].astype(str).apply(lambda x: re.sub(r'\D', '', x)).str.strip()
        combined_df['Email ID_Clean'] = combined_df['Email ID'].astype(str).str.lower().str.strip()

        seen_contacts = set()
        for idx, row in combined_df.iterrows():
            is_contact_dup = False
            phone = row['Phone Number_Clean']
            email = row['Email ID_Clean']

            if (phone and phone != 'n/a' and phone in seen_contacts) or \
               (email and email != 'n/a' and email in seen_contacts):
                is_contact_dup = True
            
            if phone and phone != 'n/a': seen_contacts.add(phone)
            if email and email != 'n/a': seen_contacts.add(email)
            
            if is_contact_dup:
                combined_df.at[idx, '_temp_status'] = 'Duplicate'
            
            # Maintain the 'Duplicate' status if it was set by the initial check for new entries
            if combined_df.at[idx, 'Status'] == 'Duplicate': 
                combined_df.at[idx, '_temp_status'] = 'Duplicate'
            
            # If not yet set, default to 'Existing' for old entries or 'New' for new entries
            if '_temp_status' not in combined_df.columns or pd.isna(combined_df.at[idx, '_temp_status']):
                combined_df.at[idx, '_temp_status'] = row['Status']


        combined_df['Status'] = combined_df['_temp_status']
        
        combined_df.drop(columns=['Phone Number_Clean', 'Email ID_Clean', '_temp_status'], errors='ignore', inplace=True)

        # Ensure consistent column names for the primary Excel output
        if 'Name' in combined_df.columns and 'Candidate Name' not in combined_df.columns:
            combined_df.rename(columns={'Name': 'Candidate Name'}, inplace=True)
        if 'Skills' in combined_df.columns and 'Skill' not in combined_df.columns:
            combined_df.rename(columns={'Skills': 'Skill'}, inplace=True)
        if 'Experience' in combined_df.columns and 'Total Experience' not in combined_df.columns:
            combined_df.rename(columns={'Experience': 'Total Experience'}, inplace=True)
        
        # Define the desired order of columns for the PRIMARY Excel file
        primary_excel_columns_order = [
            "Source Date", "Month", "Year", "Skill", "Candidate Name", 
            "Total Experience", "Email ID", "Phone Number", "File Name", "Status"
        ]
        
        # Filter and reorder columns for the primary Excel
        combined_df = combined_df[[col for col in primary_excel_columns_order if col in combined_df.columns]]

        try:
            combined_df.to_excel(excel_file_path, index=False)
            print(f"\n✅ Processing complete. Data successfully written to {excel_file_path}")
            print(f"   Total records in {excel_file_name}: {len(combined_df)}")

            # --- NEW: Generate 'cadate resume details.xlsx' ---
            cadate_excel_file_path = os.path.join(output_directory, CADATE_EXCEL_FILE_NAME)
            
            # Create a copy from the finalized combined_df
            cadate_df = combined_df.copy()
            
            # Drop 'File Name' column as requested for the second Excel
            if 'File Name' in cadate_df.columns:
                cadate_df.drop(columns=['File Name'], inplace=True)
            
            # Add new columns with default 'N/A' to cadate_df
            new_cadate_columns_to_add = [
                'Source', 'Rec', 'Education', 'NP', 'Current Company', 
                'CCTC', 'ECTC', 'Current Location', 'Current Status', 'Kishore Comment'
            ]
            for col in new_cadate_columns_to_add:
                if col not in cadate_df.columns: # Only add if it doesn't already exist from primary df
                    cadate_df[col] = 'N/A' # Default value for new columns

            # Define the exact order of columns for the SECOND Excel file
            cadate_columns_order = [
                "Source Date", "Month", "Year", "Source", "Rec", "Skill", 
                "Candidate Name", "Total Experience", "Email ID", "Phone Number", "Status", 
                "Education", "NP", "Current Company", "CCTC", "ECTC", "Current Location", 
                "Current Status", "Kishore Comment"
            ]

            # Reorder and select columns for cadate_df
            # This ensures only the requested columns are present and in the specified order.
            # Any columns from combined_df that are not in cadate_columns_order will be excluded from cadate_df.
            cadate_df = cadate_df[[col for col in cadate_columns_order if col in cadate_df.columns]]
            
            cadate_df.to_excel(cadate_excel_file_path, index=False)
            print(f"✅ Generated additional Excel: {cadate_excel_file_path} (excluding 'File Name' column).")
            print(f"   Total records in {CADATE_EXCEL_FILE_NAME}: {len(cadate_df)}")

        except Exception as e:
            print(f"❌ ERROR: Failed to write to Excel files. Primary: {excel_file_path}, Secondary: {CADATE_EXCEL_FILE_NAME}: {e}")
            traceback.print_exc() # Print full traceback for this critical error
    else:
        print("\nℹ️ No resume data extracted or added to the Excel file in this run.")

    print(f"\n🧹 Cleaning up downloaded files in: {folder_path}")
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        try:
            # Only delete files that were actually downloaded by this run
            # using the downloaded_files_info list to prevent deleting other user files.
            # Also ensures converted .docx files are cleaned up if their original .doc was processed.
            if os.path.isfile(item_path):
                # Check if it was either an original download or a converted file from an original download
                is_original_download = any(item_path == d['file_path'] for d in downloaded_files_info)
                is_converted_file = False
                for d_info in downloaded_files_info:
                    # Check if this item_path is a .docx that was converted from a .doc in d_info
                    if d_info['file_path'].endswith('.doc') and os.path.splitext(item_path)[0] == os.path.splitext(d_info['file_path'])[0] and item_path.endswith('.docx'):
                        is_converted_file = True
                        break

                if is_original_download or is_converted_file:
                    os.remove(item_path)
                    print(f"  🗑️ Deleted: {os.path.basename(item_path)}")
        except Exception as e:
            print(f"  ❌ Error deleting file {item_path}: {e}")


# --- Orchestrator for 24/7 Automation ---
def run_automation_cycle():
    print(f"\n✨✨✨ Starting Automated Resume Processing Cycle [{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ✨✨✨")
    
    downloaded_files_info = [] 
    print("\n--- Step 1: Downloading Resumes from Outlook ---")
    try:
        downloaded_files_info = download_resumes_from_outlook(
            resume_download_folder, 
            OUTLOOK_MAILBOX_NAME,
            INBOX_FOLDER,
            RESUME_KEYWORDS_IN_SUBJECT,
            RESUME_KEYWORDS_IN_BODY, 
            RESUME_KEYWORDS_IN_ATTACHMENT_NAME,
            RESUME_ATTACHMENT_EXTENSIONS
        )
        print(f"\n--- Download Summary: Downloaded {len(downloaded_files_info)} new resume(s) from Outlook. ---")
    except Exception as e:
        sanitized_error_overall = sanitize_string_for_print(str(e))
        print(f"\n❌ CRITICAL ERROR: Failed to download resumes from Outlook: {sanitized_error_overall}")
        traceback.print_exc() 
        downloaded_files_info = [] 

    print("\n--- Step 2: Processing Resumes and Updating Database ---")
    try:
        process_resumes_in_folder(resume_download_folder, output_excel_file, downloaded_files_info)
    except Exception as e:
        sanitized_error_process = sanitize_string_for_print(str(e))
        print(f"\n❌ CRITICAL ERROR: Failed to process resumes in folder: {sanitized_error_process}")
        traceback.print_exc() 

    print(f"\n✨✨✨ Automated Resume Processing Cycle Finished [{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ✨✨✨\n")


# --- Main execution block ---
if __name__ == "__main__":
    print("\n--- Initializing Resume Processor ---")
    if not os.path.exists(output_directory):
        try:
            os.makedirs(output_directory)
            print(f"  ✅ Created output directory: {output_directory}")
        except OSError as e:
            print(f"  ❌ Error creating output directory {output_directory}: {e}")
            print("     Please check directory permissions or path validity.")
            exit()
    
    if not os.path.exists(resume_download_folder): 
        try:
            os.makedirs(resume_download_folder)
            print(f"  ✅ Created resume download folder: {resume_download_folder}")
        except OSError as e:
            print(f"  ❌ Error creating resume download folder {resume_download_folder}: {e}")
            print("     Please check directory permissions or path validity.")
            exit()

    print("\n--- Starting processing cycle ---")
    run_automation_cycle()

    print("\nProcessing complete. Exiting.")