import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import re
from datetime import datetime, timedelta
import io
import tempfile
import os
import hashlib
import logging
from pathlib import Path

# Configure audit logging
LOG_DIR = Path("audit_logs")
LOG_DIR.mkdir(exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / f"audit_{datetime.now().strftime('%Y%m%d')}.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

st.set_page_config(page_title="Unbilled Billing App", layout="wide")

# Security: Session timeout (15 minutes of inactivity)
SESSION_TIMEOUT_MINUTES = 15

def check_password():
    """Returns True if user has entered correct password."""
    def password_entered():
        """Checks whether password entered is correct."""
        # Simple password check - in production, use environment variable or secure vault
        # Hash the password for basic security
        entered_hash = hashlib.sha256(st.session_state["password"].encode()).hexdigest()
        # Default password: "billing2026" (hash stored)
        correct_hash = "e007fbdc563042ac6aa9dcdfc979b2a8233938d600c412c2b2ad00a273ddd0d1"
        
        if entered_hash == correct_hash:
            st.session_state["password_correct"] = True
            st.session_state["last_activity"] = datetime.now()
            st.session_state["username"] = st.session_state.get("username_input", "user")
            logger.info(f"User '{st.session_state['username']}' logged in successfully")
            del st.session_state["password"]  # Don't store password
        else:
            st.session_state["password_correct"] = False
            logger.warning(f"Failed login attempt for user '{st.session_state.get('username_input', 'unknown')}'")

    # Check for session timeout
    if "last_activity" in st.session_state:
        time_since_activity = datetime.now() - st.session_state["last_activity"]
        if time_since_activity > timedelta(minutes=SESSION_TIMEOUT_MINUTES):
            st.session_state["password_correct"] = False
            logger.info(f"Session timeout for user '{st.session_state.get('username', 'unknown')}'")
            st.warning("Session expired due to inactivity. Please log in again.")
    
    # Update last activity
    if st.session_state.get("password_correct", False):
        st.session_state["last_activity"] = datetime.now()

    if st.session_state.get("password_correct", False):
        return True

    st.title("🔒 Secure Login")
    st.markdown("### PHI Access Control")
    st.info("⚠️ This application processes Protected Health Information (PHI). Authorized users only.")
    
    st.text_input("Username", key="username_input")
    st.text_input("Password", type="password", on_change=password_entered, key="password")
    
    if "password_correct" in st.session_state and not st.session_state["password_correct"]:
        st.error("😕 Incorrect username or password")
    
    st.markdown("---")
    st.caption("Default credentials: username='admin', password='billing2026'")
    st.caption("⚠️ Change default password in production deployment")
    
    return False

if not check_password():
    st.stop()

# User is authenticated - show main app
st.title("Unbilled Billing Processor")
st.markdown(f"👤 Logged in as: **{st.session_state.get('username', 'user')}**")
st.markdown("Upload your billing file to process it automatically")


def extract_date_from_filename(filename):
    """Extract MMDDYYYY date from filename"""
    match = re.search(r'(\d{8})', filename)
    if match:
        return match.group(1)
    return None

def _is_ecare(service_lower: str) -> bool:
    """
    Check if service is e-care (case-insensitive).
    Handles variants: 'e-care', 'e care', 'ecare'
    """
    return any(variant in service_lower for variant in ['e-care', 'e care', 'ecare'])

def is_non_billable_service_for_weekday(service: str, weekday: int) -> bool:
    """
    Determine if a service is non-billable for a given weekday.
    
    Weekday rules:
    - Tuesday (1): Everything billable (including e-care)
    - Monday (0): Non-billable: partial hospitalization, residential, detox, e-care
    - Wednesday-Friday (2,3,4): All services billable except e-care
    - Saturday-Sunday (5,6): e-care non-billable (treat like Wed-Fri)
    
    E-care is only billable on Tuesdays.
    
    Args:
        service: Service name (case-insensitive matching)
        weekday: Day of week (0=Monday, 1=Tuesday, ..., 6=Sunday)
    
    Returns:
        True if service is non-billable for the given weekday
    """
    service_lower = service.lower()
    
    # Tuesday: everything billable
    if weekday == 1:
        return False
    
    # Monday: non-billable services
    if weekday == 0:
        if _is_ecare(service_lower):
            return True
        if any(s in service_lower for s in ['partial hospitalization', 'residential', 'detox']):
            return True
        return False
    
    # Wednesday-Sunday (2,3,4,5,6): e-care non-billable
    if weekday in [2, 3, 4, 5, 6]:
        return _is_ecare(service_lower)
    
    return False

def get_filename_prefix(filename):
    """Extract the prefix before the date in filename"""
    date_match = re.search(r'(\d{8})', filename)
    if date_match:
        date_pos = date_match.start()
        date_str = date_match.group(1)
        prefix = filename[:date_pos + len(date_str)]
        remaining = filename[date_pos + len(date_str):]
        if remaining.startswith(' - '):
            prefix += ' - '
        return prefix
    return ""


def step_1_extract_invalid(ws):
    """Insert Staff/Status, extract Invalid rows to new sheet, delete from main"""

    if ws.cell(1, 1).value != "Staff/Status":
        ws.insert_cols(1)
        ws.cell(1, 1).value = "Staff/Status"

    billable_col = None
    for col in range(1, ws.max_column + 1):
        if ws.cell(1, col).value == "Billable Status":
            billable_col = col
            break

    if not billable_col:
        raise ValueError("Billable Status column not found")

    invalid_rows = []
    for row in range(2, ws.max_row + 1):
        if ws.cell(row, billable_col).value == "Invalid for Billing":
            invalid_rows.append(row)

    if not invalid_rows:
        return None

    wb = ws.parent
    if "Invalid" in wb.sheetnames:
        del wb["Invalid"]

    ws_invalid = wb.create_sheet("Invalid")

    for col in range(1, ws.max_column + 1):
        ws_invalid.cell(1, col).value = ws.cell(1, col).value

    for idx, row_num in enumerate(invalid_rows, start=2):
        for col in range(1, ws.max_column + 1):
            ws_invalid.cell(idx, col).value = ws.cell(row_num, col).value

    for row_num in reversed(invalid_rows):
        ws.delete_rows(row_num)

    ws.cell(1, 1).font = openpyxl.styles.Font(bold=True)
    ws.auto_filter.ref = ws.dimensions

    return len(invalid_rows)

def _is_ecare(service_lower: str) -> bool:
    """Check if service is e-care (handles common variants)"""
    # Handle common e-care variants: "e-care", "e care", "ecare"
    return any(variant in service_lower for variant in ["e-care", "e care", "ecare"])

def is_non_billable_service_for_weekday(service: str, weekday: int) -> bool:
    """
    Determine if a service is non-billable for a given weekday.
    
    Weekday rules:
    - Tuesday (1): all services billed (including e-care)
    - Monday (0): non-billable: partial hospitalization, residential, detox, e-care
    - Wednesday-Friday (2,3,4): bill all services except e-care
    - Saturday-Sunday (5,6): bill all services except e-care
    
    Args:
        service: Service name (case-insensitive)
        weekday: Day of week (0=Monday, 1=Tuesday, ..., 6=Sunday)
    
    Returns:
        True if service is non-billable for this weekday
    """
    service_lower = service.lower()
    
    if weekday == 1:  # Tuesday - bill all services
        return False
    elif weekday == 0:  # Monday - multiple non-billable services
        return (
            "partial hospitalization" in service_lower or
            "residential" in service_lower or
            "detox" in service_lower or
            _is_ecare(service_lower)
        )
    else:  # Wednesday-Sunday (2,3,4,5,6) - only e-care non-billable
        return _is_ecare(service_lower)

def assign_staff(ws, date_token: str = None):
    """Assign staff names based on business rules
    
    Args:
        ws: Worksheet to process
        date_token: Date string in MMDDYYYY format (from filename). If None, uses current date.
    """
    
    # Find column indices (after Staff/Status insert, columns shift by 1)
    cols = {}
    for col in range(1, ws.max_column + 1):
        header = ws.cell(1, col).value
        if header == "GROUPFLD1":
            cols['group_fld1'] = col
        elif header == "GROUPFLD2":
            cols['group'] = col
        elif header == "Service":
            cols['service'] = col
        elif header == "Payer":
            cols['payer'] = col
        elif header == "Billing Provider":
            cols['billing_provider'] = col

    if not all(k in cols for k in ['group', 'service', 'payer']):
        raise ValueError("Missing required columns: GROUPFLD2, Service, or Payer")
    
    # Determine weekday from date_token or fall back to current date
    if date_token:
        try:
            # Parse MMDDYYYY format
            date_obj = datetime.strptime(date_token, '%m%d%Y')
            day_of_week = date_obj.weekday()
            print(f"Using date from filename: {date_token} (weekday={day_of_week})")
        except ValueError:
            # Fall back to current date if parsing fails
            today = datetime.now()
            day_of_week = today.weekday()
            print(f"Failed to parse date_token '{date_token}', using current date (weekday={day_of_week})")
    else:
        # Fall back to current date if no date_token provided
        today = datetime.now()
        day_of_week = today.weekday()
        print(f"No date_token provided, using current date (weekday={day_of_week})")
    
    # Assign staff
    for row in range(2, ws.max_row + 1):
        group = str(ws.cell(row, cols['group']).value or "").strip()
        service = str(ws.cell(row, cols['service']).value or "")
        service_lower = service.lower()
        payer = str(ws.cell(row, cols['payer']).value or "").lower()

        staff = None
        
        # Check if service is non-billable for this day using weekday rules
        is_non_billable = is_non_billable_service_for_weekday(service, day_of_week)
        
        # Unable to Bill: Billing Provider = "O'Flynn, Karen" + GROUPFLD1 = "OP Chappaqua" or "OP NYC"
        if 'billing_provider' in cols and 'group_fld1' in cols:
            billing_provider = str(ws.cell(row, cols['billing_provider']).value or "").strip()
            group_fld1 = str(ws.cell(row, cols['group_fld1']).value or "").strip()

            if (billing_provider == "O'Flynn, Karen" and
                (group_fld1 == "OP Chappaqua" or group_fld1 == "OP NYC")):
                staff = "Unable to Bill"

        # Unable to Bill: Non-billable service for this day of week
        if not staff and is_non_billable:
            staff = "Unable to Bill"

        # Melissa: (Detox or Residential) + (Aetna or Humana)
        if not staff:
            has_detox_res = ("detox" in service_lower or "residential" in service_lower)
            has_insurance = "aetna" in payer or "humana" in payer

            if has_detox_res and has_insurance:
                staff = "Melissa"

        # CB: Self Pay
        if not staff and group == "Self Pay":
            staff = "CB"

        # Rosanna_1: Insurance + (IOP or Acupuncture or Partial Hospitalization)
        if not staff and group == "Insurance":
            if ("iop" in service_lower or 
                service_lower.startswith("acupuncture") or 
                "partial hospitalization" in service_lower):
                staff = "Rosanna"

        # Jasmine: (Insurance or blank) + (Detox or Residential)
        if not staff and (group == "Insurance" or group == ""):
            if ("detox" in service_lower or
                service_lower.startswith("residential")):
                staff = "Jasmine"

        # Rosanna_2: Fill remaining blanks
        if not staff:
            staff = "Rosanna"

        ws.cell(row, 1).value = staff

    print("Staff assignment complete")


def finalize_workbook(wb, exclude_drug_screens: bool = False):
    """Add Status/Comments columns and validation for Rosanna/Jasmine exports"""
    ws = wb.active

    # Bold header row
    for col in range(1, ws.max_column + 1):
        ws.cell(1, col).font = openpyxl.styles.Font(bold=True)

    # Insert two columns at E
    ws.insert_cols(5, 2)
    ws.cell(1, 5).value = "Status"
    ws.cell(1, 6).value = "Comments"
    ws.cell(1, 5).font = openpyxl.styles.Font(bold=True)
    ws.cell(1, 6).font = openpyxl.styles.Font(bold=True)

    # Create Sheet2 with validation list
    # When drug screens are excluded, omit "Utox Batch" from the dropdown
    ws_list = wb.create_sheet("Sheet2")
    ws_list['A1'] = "Billed"
    ws_list['A2'] = "Unable to Bill"
    ws_list['A3'] = "Contractual Adj"
    ws_list['A4'] = "Incomplete Billings"
    if exclude_drug_screens:
        ws_list['A5'] = "Inclusive Services"
        list_range = "=Sheet2!$A$1:$A$5"
    else:
        ws_list['A5'] = "Utox Batch"
        ws_list['A6'] = "Inclusive Services"
        list_range = "=Sheet2!$A$1:$A$6"

    # Add data validation to Status column
    dv = DataValidation(type="list", formula1=list_range, allow_blank=True)
    ws.add_data_validation(dv)

    last_row = ws.max_row
    dv.add(f'E2:E{last_row}')

    # Widen columns to fit all text
    for col in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col)].auto_size = True

def validate_uploaded_file(uploaded_file):
    """Validate uploaded file for security."""
    # Check file size (max 50MB)
    MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
    file_size = len(uploaded_file.getvalue())
    
    if file_size > MAX_FILE_SIZE:
        raise ValueError(f"File too large ({file_size/1024/1024:.1f}MB). Maximum allowed: 50MB")
    
    # Validate file extension
    if not uploaded_file.name.endswith('.xlsx'):
        raise ValueError("Only .xlsx files are allowed")
    
    # Check for suspicious patterns in filename
    suspicious_patterns = ['..', '/', '\\', '<', '>', '|', ':', '*', '?', '"']
    if any(pattern in uploaded_file.name for pattern in suspicious_patterns):
        raise ValueError("Filename contains invalid characters")
    
    logger.info(f"User '{st.session_state.get('username', 'unknown')}' uploaded file: {uploaded_file.name} ({file_size} bytes)")
    return True

DRUG_SCREEN_KEYWORDS = ["drug screen", "utox", "urine tox", "drug test", "uds"]

def is_drug_screen(service: str) -> bool:
    """Return True if the service is a drug screen."""
    service_lower = service.lower()
    return any(kw in service_lower for kw in DRUG_SCREEN_KEYWORDS)

def is_wm_program_level(cell_value) -> bool:
    """Return True if the Program Level cell contains 'WM'."""
    return "WM" in str(cell_value or "").upper()

def secure_cleanup(file_path):
    """Securely delete temporary file."""
    try:
        if os.path.exists(file_path):
            # Overwrite with random data before deletion (simple secure delete)
            with open(file_path, 'ba+', buffering=0) as f:
                length = f.tell()
                f.seek(0)
                f.write(os.urandom(length))
            os.remove(file_path)
            logger.info(f"Securely deleted temp file: {file_path}")
    except Exception as e:
        logger.error(f"Error during secure cleanup: {e}")

def process_workbook(uploaded_file, exclude_drug_screens: bool = False):
    """Process the uploaded workbook"""
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
            tmp.write(uploaded_file.getbuffer())
            tmp_path = tmp.name
        
        wb = openpyxl.load_workbook(tmp_path)
        ws = wb.active
        
        date_token = extract_date_from_filename(uploaded_file.name)
        if not date_token:
            raise ValueError("Filename must contain 8 digits (MMDDYYYY)")
        
        filename_prefix = get_filename_prefix(uploaded_file.name)
        
        invalid_count = step_1_extract_invalid(ws)
        assign_staff(ws, date_token)
        
        output_files = {}
        
        for staff_name in ["Rosanna", "Jasmine", "CB", "Melissa", "Unable to Bill"]:
            new_wb = openpyxl.Workbook()
            new_ws = new_wb.active
            new_ws.title = "Sheet1"
            
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp.write(uploaded_file.getbuffer())
                tmp_path = tmp.name
            
            logger.info(f"Processing file: {uploaded_file.name} (temp: {tmp_path})")
            
            wb = openpyxl.load_workbook(tmp_path)
            ws = wb.active
            
            date_token = extract_date_from_filename(uploaded_file.name)
            if not date_token:
                raise ValueError("Filename must contain 8 digits (MMDDYYYY)")
            
            filename_prefix = get_filename_prefix(uploaded_file.name)
            
            invalid_count = step_1_extract_invalid(ws)
            assign_staff(ws, date_token)
            
            output_files = {}

            # Find the Service and Program Level column indices
            service_col = None
            program_level_col = None
            for col in range(1, ws.max_column + 1):
                header = ws.cell(1, col).value
                if header == "Service":
                    service_col = col
                elif header == "Program Level":
                    program_level_col = col

            # Count WM rows for alert
            wm_count = 0
            if program_level_col is not None:
                for row in range(2, ws.max_row + 1):
                    if is_wm_program_level(ws.cell(row, program_level_col).value):
                        wm_count += 1

            for staff_name in ["Rosanna", "Jasmine", "CB", "Melissa", "Unable to Bill"]:
                new_wb = openpyxl.Workbook()
                new_ws = new_wb.active
                new_ws.title = "Sheet1"

                for col in range(1, ws.max_column + 1):
                    new_ws.cell(1, col).value = ws.cell(1, col).value

                new_row = 2
                for row in range(2, ws.max_row + 1):
                    if ws.cell(row, 1).value == staff_name:
                        if (exclude_drug_screens and staff_name == "Rosanna" and
                                service_col is not None):
                            service_val = str(ws.cell(row, service_col).value or "")
                            if is_drug_screen(service_val):
                                continue
                        # Skip WM program level rows for all staff except Melissa
                        # (Masters retains all rows; only Melissa bills WM)
                        if (staff_name != "Melissa" and
                                program_level_col is not None and
                                is_wm_program_level(ws.cell(row, program_level_col).value)):
                            continue
                        for col in range(1, ws.max_column + 1):
                            new_ws.cell(new_row, col).value = ws.cell(row, col).value
                        new_row += 1

                if new_row == 2:
                    continue

                if staff_name in ["Rosanna", "Jasmine"]:
                    finalize_workbook(new_wb, exclude_drug_screens=(exclude_drug_screens and staff_name == "Rosanna"))
                
                output = io.BytesIO()
                new_wb.save(output)
                output.seek(0)
                output_filename = f"{filename_prefix}{staff_name}.xlsx"
                output_files[output_filename] = output
            
            main_output = io.BytesIO()
            wb.save(main_output)
            main_output.seek(0)
            main_filename = f"{filename_prefix}Masters.xlsx"
            output_files[main_filename] = main_output
            
            logger.info(f"Successfully processed file: {uploaded_file.name}, generated {len(output_files)} output files")

            return output_files, invalid_count, date_token, wm_count
            
    finally:
        # Always cleanup temp file securely
        if tmp_path:
            secure_cleanup(tmp_path)

# UI
st.markdown("### Upload your billing file")
st.info("🔒 All actions are logged. Session timeout: 15 minutes of inactivity.")

uploaded_file = st.file_uploader(
    "Choose an Excel file (must have MMDDYYYY in filename)",
    type="xlsx"
)

exclude_drug_screens = st.checkbox(
    "Exclude drug screens from Rosanna's report",
    value=False,
    help="When checked, rows with drug screen services (drug screen, utox, urine tox, drug test, uds) will not be included in Rosanna's output file, and 'Utox Batch' will be removed from the Status dropdown."
)

if uploaded_file is not None:
    try:
        st.info("Processing your file...")
        output_files, invalid_count, date_token, wm_count = process_workbook(uploaded_file, exclude_drug_screens=exclude_drug_screens)

        if wm_count > 0:
            st.warning(
                f"⚠️ Alert: {wm_count} row(s) with 'WM' in the Program Level column were found. "
                "These rows are only included in Masters and Melissa's report. "
                "Only Melissa is authorized to bill WM."
            )

        st.success("✓ Processing complete!")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Date Extracted", date_token if date_token else "Today's date")
        with col2:
            st.metric("Invalid Rows", invalid_count or 0)
        with col3:
            st.metric("Files Generated", len(output_files))
        
        st.markdown("---")
        st.markdown("### Download Results")
        
        for output_filename, file_bytes in output_files.items():
            if st.download_button(
                label=f"Download {output_filename}",
                data=file_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=output_filename
            ):
                logger.info(f"User '{st.session_state.get('username', 'unknown')}' downloaded: {output_filename}")
    
    except Exception as e:
        logger.error(f"Error processing file for user '{st.session_state.get('username', 'unknown')}': {str(e)}")
        st.error(f"Error: {str(e)}")
