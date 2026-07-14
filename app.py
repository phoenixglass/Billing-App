import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import re
from datetime import datetime
import io
import tempfile
import os
import json
import logging
from pathlib import Path

from billing_rules import (
    parse_weekday_from_token,
    is_non_billable_service_for_weekday,
    is_professional_claim_type,
    is_iop_service,
    is_php_service,
    _is_drug_screen as is_drug_screen,
    ROSANNA_PROFESSIONAL_CAP,
    JOSHUA_PROFESSIONAL_CAP,
)

try:
    from google.cloud import storage
    from google.oauth2 import service_account
    GCS_AVAILABLE = True
except ImportError:
    GCS_AVAILABLE = False

# Configure audit logging
LOG_DIR = Path("audit_logs")
LOG_DIR.mkdir(exist_ok=True)

def setup_logging():
    """Setup logging with GCS support if credentials are available."""
    handlers = []
    gcs_handler = None

    # Try to setup GCS logging (for Streamlit Cloud)
    if GCS_AVAILABLE and "gcp_service_account" in st.secrets:
        try:
            creds_json = st.secrets["gcp_service_account"]
            if isinstance(creds_json, str):
                creds_dict = json.loads(creds_json)
            else:
                creds_dict = dict(creds_json)

            credentials = service_account.Credentials.from_service_account_info(creds_dict)
            gcs_client = storage.Client(credentials=credentials)
            bucket_name = st.secrets.get("gcs_bucket_name", "billing-app-audit-logs")
            bucket = gcs_client.bucket(bucket_name)

            # Verify bucket exists
            if bucket.exists():
                gcs_handler = GCSHandler(gcs_client, bucket_name)
                handlers.append(gcs_handler)
            else:
                print(f"WARNING: GCS bucket '{bucket_name}' not found. Using local logging only.")
        except Exception as e:
            print(f"WARNING: Failed to setup GCS logging: {e}. Using local logging only.")

    # Always add local logging handler for development/fallback
    log_file = LOG_DIR / f"audit_{datetime.now().strftime('%Y%m%d')}.log"
    handlers.append(logging.FileHandler(log_file))

    # Add console handler
    handlers.append(logging.StreamHandler())

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=handlers,
        force=True
    )

    # Security: Enforce restrictive permissions on audit logs (owner read/write only)
    for log_file in LOG_DIR.glob("audit_*.log"):
        os.chmod(log_file, 0o600)

    return gcs_handler

class GCSHandler(logging.Handler):
    """Custom logging handler that writes to Google Cloud Storage."""

    def __init__(self, gcs_client, bucket_name):
        super().__init__()
        self.gcs_client = gcs_client
        self.bucket_name = bucket_name
        self.bucket = gcs_client.bucket(bucket_name)

    def emit(self, record):
        """Write log record to GCS."""
        try:
            log_entry = self.format(record) + "\n"
            log_date = datetime.now().strftime('%Y%m%d')
            blob_name = f"audit_logs/audit_{log_date}.log"

            blob = self.bucket.blob(blob_name)

            # Append to existing log
            if blob.exists():
                existing = blob.download_as_string().decode('utf-8')
                blob.upload_from_string(existing + log_entry)
            else:
                blob.upload_from_string(log_entry)
        except Exception as e:
            self.handleError(record)

gcs_handler = setup_logging()
logger = logging.getLogger(__name__)

if gcs_handler:
    logger.info("✓ GCS logging enabled")

st.set_page_config(page_title="Unbilled Billing App", layout="wide")

st.title("Unbilled Billing Processor")
st.markdown("Upload your billing file to process it automatically")


def extract_date_from_filename(filename):
    """Extract MMDDYYYY date from filename"""
    match = re.search(r'(\d{8})', filename)
    if match:
        return match.group(1)
    return None

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

    if invalid_rows:
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

    for col in range(1, ws.max_column + 1):
        ws.cell(1, col).font = openpyxl.styles.Font(bold=True)
    ws.auto_filter.ref = ws.dimensions

    return len(invalid_rows) if invalid_rows else 0

def assign_staff(ws, date_token: str = None):
    """Assign staff names based on the standard daily billing rules.

    - Self Pay (GROUPFLD2 == "Self Pay") always goes to CB; every service
      bills every day, with no exceptions.
    - Jasmine, Rosanna, and Joshua only ever receive GROUPFLD2 ==
      "Insurance" rows.
    - Among Insurance rows, the "professional pool" is every row whose
      Claim Type is CMS-1500, sorted alphabetically by Client. On Monday,
      Rosanna receives the first ROSANNA_PROFESSIONAL_CAP[weekday] of that
      sorted pool; on Tuesday/Thursday/Friday, Joshua receives the first
      JOSHUA_PROFESSIONAL_CAP[weekday] instead (0 on Wednesdays/weekends,
      so nobody caps the pool those days); the rest of the pool goes to
      Jasmine.
    - Jasmine also receives Insurance rows that are billable Programming
      (Detox, Residential, IOP) or e-care for that weekday.
    - IOP (including Telemed IOP) always goes to Jasmine when billable that
      weekday, bypassing the professional pool/Rosanna/Joshua split even if
      Claim Type is CMS-1500.
    - Any other Insurance row (not billable that day), or any row that is
      neither Self Pay nor Insurance, is Unable to Bill.
    - Melissa (WM/OP WM Program Level, PHP/Partial Hospitalization, or
      Aetna/Humana Detox/Residential) and the O'Flynn Karen OP
      Chappaqua/OP NYC "Unable to Bill" rule take priority over all of the
      above. PHP rows are assigned to Melissa every day in the Masters
      spreadsheet; she does not get an individual report, and PHP is
      billed only on Tuesdays as an operational matter.

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
        elif header == "Program Level":
            cols['program_level'] = col
        elif header == "Client":
            cols['client'] = col
        elif header == "Claim Type":
            cols['claim_type'] = col

    if not all(k in cols for k in ['group', 'service', 'payer', 'claim_type', 'client']):
        raise ValueError(
            "Missing required columns: GROUPFLD2, Service, Payer, Claim Type, or Client"
        )

    weekday, did_fallback = parse_weekday_from_token(date_token)
    if did_fallback:
        print(f"Failed to parse date_token '{date_token}', using current weekday={weekday}")
    else:
        print(f"Using date from filename: {date_token} (weekday={weekday})")

    rosanna_cap = ROSANNA_PROFESSIONAL_CAP.get(weekday, 0)
    joshua_cap = JOSHUA_PROFESSIONAL_CAP.get(weekday, 0)
    if rosanna_cap:
        capped_staff, professional_cap = "Rosanna", rosanna_cap
    elif joshua_cap:
        capped_staff, professional_cap = "Joshua", joshua_cap
    else:
        capped_staff, professional_cap = None, 0

    row_data_map = {}
    fixed_staff = {}        # original_row -> staff already decided
    professional_rows = []  # (original_row, client) still needing the capped-staff/Jasmine split
    other_rows = []         # original_row order for every row not in the professional pool

    for row in range(2, ws.max_row + 1):
        row_data_map[row] = [ws.cell(row, col).value for col in range(1, ws.max_column + 1)]

        group = str(ws.cell(row, cols['group']).value or "").strip()
        service = str(ws.cell(row, cols['service']).value or "")
        payer = str(ws.cell(row, cols['payer']).value or "").lower()
        claim_type = str(ws.cell(row, cols['claim_type']).value or "")

        # CB: Self Pay bills every service every day, no exceptions. This
        # overrides every other rule.
        if group == "Self Pay":
            fixed_staff[row] = "CB"
            other_rows.append(row)
            continue

        staff = None

        # WM / OP WM Program Level → Melissa
        if 'program_level' in cols and is_wm_program_level(ws.cell(row, cols['program_level']).value):
            staff = "Melissa"

        # Unable to Bill: Billing Provider = "O'Flynn, Karen" + GROUPFLD1 = "OP Chappaqua" or "OP NYC"
        if not staff and 'billing_provider' in cols and 'group_fld1' in cols:
            billing_provider = str(ws.cell(row, cols['billing_provider']).value or "").strip()
            group_fld1 = str(ws.cell(row, cols['group_fld1']).value or "").strip()
            if (billing_provider == "O'Flynn, Karen" and
                    group_fld1 in ("OP Chappaqua", "OP NYC")):
                staff = "Unable to Bill"

        # Melissa: (Detox or Residential) + (Aetna or Humana), but not drug screens
        if not staff:
            service_lower = service.lower()
            has_detox_res = ("detox" in service_lower or "residential" in service_lower)
            has_insurance = "aetna" in payer or "humana" in payer
            if has_detox_res and has_insurance and not is_drug_screen(service):
                staff = "Melissa"

        # PHP/Partial Hospitalization always goes to Melissa, every day.
        if not staff and is_php_service(service):
            staff = "Melissa"

        if staff:
            fixed_staff[row] = staff
            other_rows.append(row)
            continue

        # Jasmine, Rosanna, and Joshua only ever receive Insurance rows.
        if group != "Insurance":
            fixed_staff[row] = "Unable to Bill"
            other_rows.append(row)
            continue

        # IOP (including Telemed IOP) always goes to Jasmine when billable
        # that weekday, bypassing the professional pool/Rosanna/Joshua split
        # even if Claim Type is CMS-1500.
        if is_iop_service(service):
            if is_non_billable_service_for_weekday(service, weekday):
                fixed_staff[row] = "Unable to Bill"
            else:
                fixed_staff[row] = "Jasmine"
            other_rows.append(row)
            continue

        if is_professional_claim_type(claim_type):
            client = str(ws.cell(row, cols['client']).value or "").strip()
            professional_rows.append((row, client))
            continue

        if is_non_billable_service_for_weekday(service, weekday):
            fixed_staff[row] = "Unable to Bill"
        else:
            fixed_staff[row] = "Jasmine"
        other_rows.append(row)

    # Move the professional (Insurance + CMS-1500) pool to the top of the
    # sheet, sorted alphabetically by Client; every other row keeps its
    # original relative order after that.
    professional_rows.sort(key=lambda x: x[1].lower())

    ordered_rows = [row for row, _ in professional_rows] + other_rows
    new_row_pos = 2
    for original_row in ordered_rows:
        for col in range(1, ws.max_column + 1):
            ws.cell(new_row_pos, col).value = row_data_map[original_row][col - 1]
        new_row_pos += 1

    num_capped = min(professional_cap, len(professional_rows))
    new_row_pos = 2
    for idx in range(len(professional_rows)):
        ws.cell(new_row_pos, 1).value = capped_staff if (capped_staff and idx < num_capped) else "Jasmine"
        new_row_pos += 1
    for original_row in other_rows:
        ws.cell(new_row_pos, 1).value = fixed_staff[original_row]
        new_row_pos += 1

    print("Staff assignment complete")


def finalize_workbook(wb, include_batch_billings: bool = False, include_iop_status: bool = False,
                      skip_status_columns: bool = False):
    """Add Status/Comments columns and validation for Rosanna/Joshua/Jasmine exports.

    Args:
        wb: Workbook to finalize.
        include_batch_billings: When True, add 'Batch Billings' to the dropdown (Jasmine only).
        include_iop_status: When True, add 'IOP' to the dropdown (Jasmine only).
        skip_status_columns: When True, skip adding Status/Comments columns (for CB/self-pay).
    """
    ws = wb.active

    # Bold header row
    for col in range(1, ws.max_column + 1):
        ws.cell(1, col).font = openpyxl.styles.Font(bold=True)

    # Insert two columns at E (unless skipped for CB/self-pay)
    if not skip_status_columns:
        ws.insert_cols(5, 2)
        ws.cell(1, 5).value = "Status"
        ws.cell(1, 6).value = "Comments"
        ws.cell(1, 5).font = openpyxl.styles.Font(bold=True)
        ws.cell(1, 6).font = openpyxl.styles.Font(bold=True)

        # Create Sheet2 with validation list
        ws_list = wb.create_sheet("Sheet2")
        status_items = ["Billed", "Unable to Bill", "Contractual Adj", "Incomplete Billings",
                        "Utox Batch", "Inclusive Services"]
        if include_batch_billings:
            status_items.append("Batch Billings")
        if include_iop_status:
            status_items.append("IOP")

        for idx, item in enumerate(status_items, start=1):
            ws_list[f'A{idx}'] = item
        list_range = f"=Sheet2!$A$1:$A${len(status_items)}"

        # Add data validation to Status column
        dv = DataValidation(type="list", formula1=list_range, allow_blank=True)
        ws.add_data_validation(dv)

        last_row = ws.max_row
        dv.add(f'E2:E{last_row}')

    # Widen columns to fit all text. openpyxl's auto_size flag is unreliable in
    # Excel, so compute an explicit width from the longest value in each column.
    for col in range(1, ws.max_column + 1):
        max_length = 0
        for row in range(1, ws.max_row + 1):
            value = ws.cell(row, col).value
            if value is None:
                continue
            length = len(str(value))
            if length > max_length:
                max_length = length
        # Add a small padding and cap the width so a single huge cell doesn't
        # blow out the layout.
        ws.column_dimensions[get_column_letter(col)].width = min(max_length + 2, 60)

    # Enable filtering on the header row.
    ws.auto_filter.ref = ws.dimensions

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
    
    logger.info(f"Uploaded file: {uploaded_file.name} ({file_size} bytes)")
    return True

def is_wm_program_level(cell_value) -> bool:
    """Return True if the Program Level cell contains 'WM'."""
    return "WM" in str(cell_value or "").upper()

def is_op_wm_program_level(cell_value) -> bool:
    """Return True if the Program Level cell contains 'OP WM'."""
    return "OP WM" in str(cell_value or "").upper()

def is_anthem_payer(payer: str) -> bool:
    """Return True if the payer contains 'anthem'."""
    return "anthem" in payer.lower()

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

def is_bcb_anthem_ct_php_res_detox(payer: str, service: str) -> bool:
    """Return True if payer is BCB Anthem CT and service is PHP, Residential, or Detox."""
    return ("bcb anthem ct" in payer.lower() and
            any(s in service.lower() for s in ["partial hospitalization", "residential", "detox"]))

def process_workbook(uploaded_file, exclude_optum: bool = False,
                     exclude_bcb_anthem_ct: bool = False,
                     exclude_anthem_rosanna_jasmine_owm: bool = False,
                     exclude_detox_residential: bool = False):
    """Process the uploaded workbook"""
    tmp_path = None
    try:
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

        # Find the Service, Program Level, and Payer column indices
        service_col = None
        program_level_col = None
        payer_col = None
        for col in range(1, ws.max_column + 1):
            header = ws.cell(1, col).value
            if header == "Service":
                service_col = col
            elif header == "Program Level":
                program_level_col = col
            elif header == "Payer":
                payer_col = col

        # Count WM rows for alerts (both WM and OP WM go to Melissa)
        wm_count = 0
        if program_level_col is not None:
            for row in range(2, ws.max_row + 1):
                pl = ws.cell(row, program_level_col).value
                if is_wm_program_level(pl):
                    wm_count += 1

        # Rosanna, Joshua, Jasmine, and CB always get individual reports
        # (empty ones are skipped below). Melissa, Unable to Bill, etc. stay
        # Masters-only.
        for staff_name in ["Rosanna", "Joshua", "Jasmine", "CB"]:
            new_wb = openpyxl.Workbook()
            new_ws = new_wb.active
            new_ws.title = "Sheet1"

            for col in range(1, ws.max_column + 1):
                new_ws.cell(1, col).value = ws.cell(1, col).value

            new_row = 2
            for row in range(2, ws.max_row + 1):
                assigned_staff = ws.cell(row, 1).value
                if assigned_staff == staff_name:
                    # Exclude Optum utox rows from all individual workbooks
                    if (exclude_optum and service_col is not None and
                            payer_col is not None):
                        service_val = str(ws.cell(row, service_col).value or "")
                        payer_val = str(ws.cell(row, payer_col).value or "").lower()
                        if is_drug_screen(service_val) and "optum" in payer_val:
                            continue
                    if (exclude_bcb_anthem_ct and service_col is not None and
                            payer_col is not None):
                        service_val = str(ws.cell(row, service_col).value or "")
                        payer_val = str(ws.cell(row, payer_col).value or "")
                        if is_bcb_anthem_ct_php_res_detox(payer_val, service_val):
                            continue
                    if (exclude_anthem_rosanna_jasmine_owm and
                            assigned_staff in ("Rosanna", "Jasmine") and
                            payer_col is not None):
                        payer_val = str(ws.cell(row, payer_col).value or "")
                        if is_anthem_payer(payer_val):
                            continue
                    if exclude_detox_residential and service_col is not None:
                        service_val = str(ws.cell(row, service_col).value or "").lower()
                        if "detox" in service_val or "residential" in service_val:
                            continue
                    # Skip WM/OP WM program level rows for all staff except Melissa
                    if (staff_name != "Melissa" and
                            program_level_col is not None and
                            is_wm_program_level(ws.cell(row, program_level_col).value)):
                        continue
                    for col in range(1, ws.max_column + 1):
                        new_ws.cell(new_row, col).value = ws.cell(row, col).value
                    new_row += 1

            if new_row == 2:
                continue

            if staff_name in ("Rosanna", "Joshua", "Jasmine"):
                finalize_workbook(new_wb, include_batch_billings=(staff_name == "Jasmine"),
                                  include_iop_status=(staff_name == "Jasmine"))
            elif staff_name == "CB":
                finalize_workbook(new_wb, skip_status_columns=True)

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
    type="xlsx",
    key="billing_file_uploader"
)

# Preserve uploaded file in session state to prevent it from disappearing on reruns
if uploaded_file is not None:
    st.session_state["current_uploaded_file"] = uploaded_file

st.markdown(
    "Staff assignment now follows the standard daily schedule automatically, based on "
    "the date in the filename: Self Pay always bills every day to **CB**; for Insurance "
    "rows, **Rosanna** (Mondays) or **Joshua** (Tuesdays, Thursdays, Fridays) gets the "
    "day's capped number of Professional (Claim Type CMS-1500) services sorted "
    "alphabetically by Client, and **Jasmine** gets the remaining Professional rows plus "
    "any billable Programming/e-care rows for that day. PHP rows always go to **Melissa** "
    "in the Masters report (no individual report is generated for her)."
)

exclude_optum = st.checkbox(
    "Exclude Optum insurance",
    value=False,
    help="When checked, utox (drug screen) rows with Optum as the payer will be excluded from all individual staff workbooks."
)

exclude_bcb_anthem_ct = st.checkbox(
    "Exclude BCB Anthem CT for PHP, Residential, and Detox",
    value=False,
    help="When checked, rows where the payer is BCB Anthem CT and the service is PHP (Partial Hospitalization), Residential, or Detox will be excluded from all individual staff workbooks."
)

exclude_anthem_rosanna_jasmine_owm = st.checkbox(
    "Remove Anthem from Rosanna and Jasmine reports",
    value=False,
    help="When checked, rows where the payer contains 'Anthem' will be excluded from Rosanna's and Jasmine's workbooks. Anthem rows are still retained in the Masters report."
)

exclude_detox_residential = st.checkbox(
    "Don't give anyone Detox or Residential",
    value=False,
    help="When checked, Detox and Residential service rows are excluded from all individual staff workbooks. They still appear in the Masters report."
)

if uploaded_file is not None:
    try:
        validate_uploaded_file(uploaded_file)
        st.info("Processing your file...")
        output_files, invalid_count, date_token, wm_count = process_workbook(
            uploaded_file,
            exclude_optum=exclude_optum,
            exclude_bcb_anthem_ct=exclude_bcb_anthem_ct,
            exclude_anthem_rosanna_jasmine_owm=exclude_anthem_rosanna_jasmine_owm,
            exclude_detox_residential=exclude_detox_residential,
        )

        if wm_count > 0:
            st.warning(
                f"⚠️ Alert: {wm_count} row(s) with 'WM' or 'OP WM' in the Program Level column were found. "
                "These rows are assigned to Melissa and appear only in the Masters report. "
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
                logger.info(f"Downloaded: {output_filename}")

    except Exception as e:
        logger.error(f"Error processing file: {str(e)}")
        st.error(f"Error: {str(e)}")
