import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import re
from datetime import datetime, timedelta
import io
import tempfile
import os
import json
import hmac
import hashlib
import secrets
import logging
from pathlib import Path

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

# Security: Session timeout (15 minutes of inactivity)
SESSION_TIMEOUT_MINUTES = 15

# Security: brute-force lockout after repeated failures
MAX_FAILED_ATTEMPTS = 5
LOCKOUT_MINUTES = 15

# PBKDF2 work factor for password hashing. Tune upward as hardware improves.
PBKDF2_ITERATIONS = 600_000

# Environment variable holding a JSON object of {username: pbkdf2_hash}.
# Generate a hash with hash_password() (see SECURITY.md). No default/built-in
# credentials are shipped: if this is unset the app refuses to authenticate.
CREDENTIALS_ENV_VAR = "BILLING_APP_CREDENTIALS"


def hash_password(password: str, *, iterations: int = PBKDF2_ITERATIONS) -> str:
    """Return a salted PBKDF2-HMAC-SHA256 hash string for storage."""
    salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations)
    return f"pbkdf2_sha256${iterations}${salt.hex()}${dk.hex()}"


def verify_password(password: str, stored: str) -> bool:
    """Constant-time verification of a password against a stored PBKDF2 hash."""
    try:
        algo, iter_s, salt_hex, hash_hex = stored.split("$")
        if algo != "pbkdf2_sha256":
            return False
        dk = hashlib.pbkdf2_hmac(
            "sha256", password.encode("utf-8"), bytes.fromhex(salt_hex), int(iter_s)
        )
        return hmac.compare_digest(dk.hex(), hash_hex)
    except (ValueError, AttributeError):
        return False


def load_credentials() -> dict:
    """Load {username: pbkdf2_hash} from the credentials environment variable."""
    raw = os.environ.get(CREDENTIALS_ENV_VAR)
    if not raw:
        return {}
    try:
        creds = json.loads(raw)
    except json.JSONDecodeError:
        logger.error("%s is set but is not valid JSON", CREDENTIALS_ENV_VAR)
        return {}
    if not isinstance(creds, dict):
        logger.error("%s must be a JSON object of {username: hash}", CREDENTIALS_ENV_VAR)
        return {}
    return creds


def check_password():
    """Returns True if the user is authenticated."""
    credentials = load_credentials()
    if not credentials:
        st.title("🔒 Secure Login")
        st.error(
            f"Authentication is not configured. Set the {CREDENTIALS_ENV_VAR} "
            "environment variable before starting the app (see SECURITY.md)."
        )
        logger.error("Login blocked: %s is not configured", CREDENTIALS_ENV_VAR)
        return False

    def password_entered():
        """Validate the submitted username/password pair."""
        username = st.session_state.get("username_input", "").strip()
        password = st.session_state.get("password", "")

        # Enforce lockout window before checking the credential.
        locked_until = st.session_state.get("lockout_until")
        if locked_until and datetime.now() < locked_until:
            st.session_state["password_correct"] = False
            logger.warning("Login attempt during active lockout for user '%s'", username or "unknown")
            return

        stored = credentials.get(username)
        if stored and verify_password(password, stored):
            st.session_state["password_correct"] = True
            st.session_state["last_activity"] = datetime.now()
            st.session_state["username"] = username
            st.session_state["failed_attempts"] = 0
            st.session_state.pop("lockout_until", None)
            logger.info("User '%s' logged in successfully", username)
        else:
            st.session_state["password_correct"] = False
            attempts = st.session_state.get("failed_attempts", 0) + 1
            st.session_state["failed_attempts"] = attempts
            logger.warning("Failed login attempt for user '%s' (attempt %d)", username or "unknown", attempts)
            if attempts >= MAX_FAILED_ATTEMPTS:
                st.session_state["lockout_until"] = datetime.now() + timedelta(minutes=LOCKOUT_MINUTES)
                logger.warning("User '%s' locked out for %d minutes", username or "unknown", LOCKOUT_MINUTES)

        # Never retain the submitted password in session state.
        st.session_state.pop("password", None)

    # Check for session timeout
    if "last_activity" in st.session_state:
        time_since_activity = datetime.now() - st.session_state["last_activity"]
        if time_since_activity > timedelta(minutes=SESSION_TIMEOUT_MINUTES):
            st.session_state["password_correct"] = False
            logger.info("Session timeout for user '%s'", st.session_state.get("username", "unknown"))
            st.warning("Session expired due to inactivity. Please log in again.")

    # Update last activity
    if st.session_state.get("password_correct", False):
        st.session_state["last_activity"] = datetime.now()
        return True

    st.title("🔒 Secure Login")
    st.markdown("### PHI Access Control")
    st.info("⚠️ This application processes Protected Health Information (PHI). Authorized users only.")

    locked_until = st.session_state.get("lockout_until")
    if locked_until and datetime.now() < locked_until:
        remaining = int((locked_until - datetime.now()).total_seconds() // 60) + 1
        st.error(f"Too many failed attempts. Try again in ~{remaining} minute(s).")
        return False

    st.text_input("Username", key="username_input")
    st.text_input("Password", type="password", on_change=password_entered, key="password")

    if "password_correct" in st.session_state and not st.session_state["password_correct"]:
        st.error("😕 Incorrect username or password")

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

def _is_ecare(service: str) -> bool:
    """
    Check if service is e-care (case-insensitive).
    Handles variants: 'e-care', 'e care', 'ecare', 'extended care'
    """
    s = service.lower()
    return any(variant in s for variant in ['e-care', 'e care', 'ecare', 'extended care'])

def _is_programming_service(service: str) -> bool:
    """Check if service is a Programming service: Detox, Residential, PHP, or IOP."""
    s = service.lower()
    return (
        'detox' in s
        or 'residential' in s
        or 'partial hospitalization' in s
        or 'php' in s
        or 'iop' in s
    )

def is_non_billable_service_for_weekday(
    service: str, weekday: int, php_on_monday: bool = False
) -> bool:
    """
    Determine if a service is non-billable for a given weekday.

    Weekly billing schedule:
    - Monday (0):    Professional + Utox     (Programming + e-care non-billable)
    - Tuesday (1):   E-care + Programming    (Professional + Utox non-billable)
    - Wednesday (2): Professional + Utox     (Programming + e-care non-billable)
    - Thursday (3):  Programming + Professional + Utox (e-care non-billable)
    - Friday (4):    Programming + Professional + Utox (e-care non-billable)
    - Saturday/Sunday (5,6): all services billable except e-care

    Service categories:
    - Programming: Detox, Residential, Partial Hospitalization (PHP), IOP
    - Professional: all other services (including Acupuncture)
    - Utox: drug screen services

    E-care is billable on Tuesdays only.

    Args:
        service: Service name (case-insensitive matching)
        weekday: Day of week (0=Monday, 1=Tuesday, ..., 6=Sunday)
        php_on_monday: When True, Partial Hospitalization is billable on Mondays

    Returns:
        True if service is non-billable for the given weekday
    """
    service_lower = service.lower()

    # e-care is billable on Tuesdays only
    if _is_ecare(service_lower):
        return weekday != 1

    # Drug screens (Utox) follow Professional days: billable Mon/Wed/Thu/Fri
    # and weekends; non-billable only on Tuesday.
    if is_drug_screen(service):
        return weekday == 1

    is_programming = _is_programming_service(service_lower)

    # Monday and Wednesday: Professional + Utox only (Programming non-billable)
    if weekday in (0, 2):
        if not is_programming:
            return False
        if weekday == 0 and php_on_monday and (
            'partial hospitalization' in service_lower or 'php' in service_lower
        ):
            return False
        return True

    # Tuesday: E-care + Programming only (Professional + Utox non-billable)
    if weekday == 1:
        return not is_programming

    # Thursday and Friday: Programming + Professional + Utox all billable
    if weekday in (3, 4):
        return False

    # Saturday and Sunday: everything billable (e-care handled above)
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

def _is_ecare(service: str) -> bool:
    """Check if service is e-care (handles common variants)"""
    # Handle common e-care variants: "e-care", "e care", "ecare", "extended care"
    s = service.lower()
    return any(variant in s for variant in ["e-care", "e care", "ecare", "extended care"])

def _is_programming_service(service: str) -> bool:
    """Check if service is a Programming service: Detox, Residential, PHP, or IOP."""
    s = service.lower()
    return (
        "detox" in s
        or "residential" in s
        or "partial hospitalization" in s
        or "php" in s
        or "iop" in s
    )

def is_non_billable_service_for_weekday(
    service: str, weekday: int, php_on_monday: bool = False
) -> bool:
    """
    Determine if a service is non-billable for a given weekday.

    Weekly billing schedule:
    - Monday (0):    Professional + Utox     (Programming + e-care non-billable)
    - Tuesday (1):   E-care + Programming    (Professional + Utox non-billable)
    - Wednesday (2): Professional + Utox     (Programming + e-care non-billable)
    - Thursday (3):  Programming + Professional + Utox (e-care non-billable)
    - Friday (4):    Programming + Professional + Utox (e-care non-billable)
    - Saturday/Sunday (5,6): all services billable except e-care

    Service categories:
    - Programming: Detox, Residential, Partial Hospitalization (PHP), IOP
    - Professional: all other services (including Acupuncture)
    - Utox: drug screen services

    E-care is billable on Tuesdays only.

    Args:
        service: Service name (case-insensitive)
        weekday: Day of week (0=Monday, 1=Tuesday, ..., 6=Sunday)
        php_on_monday: When True, Partial Hospitalization is billable on Mondays

    Returns:
        True if service is non-billable for this weekday
    """
    service_lower = service.lower()

    # e-care is billable on Tuesdays only
    if _is_ecare(service_lower):
        return weekday != 1

    # Drug screens (Utox) follow Professional days: billable Mon/Wed/Thu/Fri
    # and weekends; non-billable only on Tuesday.
    if is_drug_screen(service):
        return weekday == 1

    is_programming = _is_programming_service(service_lower)

    # Monday and Wednesday: Professional + Utox only (Programming non-billable)
    if weekday in (0, 2):
        if not is_programming:
            return False
        if weekday == 0 and php_on_monday and (
            "partial hospitalization" in service_lower or "php" in service_lower
        ):
            return False
        return True

    # Tuesday: E-care + Programming only (Professional + Utox non-billable)
    if weekday == 1:
        return not is_programming

    # Thursday and Friday: Programming + Professional + Utox all billable
    if weekday in (3, 4):
        return False

    # Saturday and Sunday: everything billable (e-care handled above)
    return False

# Default redirect: Rosanna's caseload folds into Jasmine. This applies unless
# one of the explicit Rosanna routing options is checked, in which case Rosanna
# gets her own assignments and report (see process_workbook).
STAFF_REDIRECT = {"Rosanna": "Jasmine"}


def assign_staff(ws, date_token: str = None, give_utox_to_jasmine: bool = False,
                 route_iop_acu_to_rosanna: bool = False,
                 rosanna_php_iop_only: bool = False,
                 jasmine_detox_residential_only: bool = False,
                 rosanna_iop_php_acu: bool = False,
                 jasmine_inpatient_professional: bool = False,
                 php_on_monday: bool = False,
                 rosanna_iop_jasmine_php: bool = False,
                 jasmine_iop_professional: bool = False,
                 jasmine_detox_residential_php: bool = False,
                 split_professional_utox: bool = False):
    """Assign staff names based on business rules

    Args:
        ws: Worksheet to process
        date_token: Date string in MMDDYYYY format (from filename). If None, uses current date.
        give_utox_to_jasmine: If True, drug screen (utox) rows are assigned to Jasmine.
        route_iop_acu_to_rosanna: If True, only IOP and Acupuncture go to Rosanna and all
            remaining services default to Jasmine instead of Rosanna.
        rosanna_php_iop_only: If True, only PHP (Partial Hospitalization) and IOP go to
            Rosanna (Acupuncture excluded).
        jasmine_detox_residential_only: If True, Jasmine only receives Detox and Residential
            rows; she is not used as the fallback default for unmatched services.
        rosanna_iop_php_acu: If True, Rosanna receives IOP, PHP, and Acupuncture services.
        jasmine_inpatient_professional: If True, Jasmine receives inpatient (Detox/Residential)
            and all professional outpatient services.
        php_on_monday: If True, Partial Hospitalization is treated as billable on Mondays.
        rosanna_iop_jasmine_php: If True, IOP goes to Rosanna and PHP (Partial Hospitalization)
            goes to Jasmine. Day-of-week rules still apply.
        jasmine_iop_professional: If True, Jasmine receives IOP, PHP, Acupuncture, and all
            professional outpatient services. Rosanna does not receive IOP, PHP, or
            Acupuncture under this option.
        jasmine_detox_residential_php: If True, Detox, Residential, and PHP (Partial
            Hospitalization) are treated as billable on every weekday and routed to
            Jasmine. Jasmine receives only these three services; all other services
            default to Rosanna.
        split_professional_utox: If True, Professional and Utox services are sorted by
            Client (column C) alphabetically. First 300 rows go to Rosanna, remaining go
            to Jasmine. All IOP goes to Jasmine.
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

    if not all(k in cols for k in ['group', 'service', 'payer']):
        raise ValueError("Missing required columns: GROUPFLD2, Service, or Payer")

    # Pre-process for split_professional_utox: collect, sort, and assign split
    split_assignment = {}
    if split_professional_utox:
        professional_utox_rows = []
        other_rows = []
        row_data_map = {}

        # Collect all row data and identify professional/utox rows
        for row in range(2, ws.max_row + 1):
            row_data = [ws.cell(row, col).value for col in range(1, ws.max_column + 1)]
            row_data_map[row] = row_data

            service = str(ws.cell(row, cols['service']).value or "")
            service_lower = service.lower()
            group = str(ws.cell(row, cols['group']).value or "").strip()

            # Collect Professional services (not IOP, not programming, not drug screens already handled)
            # Exclude Self Pay (CB) rows from the split
            is_professional = is_professional_service(service)
            is_utox = is_drug_screen(service)
            is_iop = "iop" in service_lower

            if (is_professional or is_utox) and not is_iop and group != "Self Pay":
                client = str(ws.cell(row, cols.get('client', 3)).value or "").strip()
                professional_utox_rows.append((row, client, service_lower))
            else:
                other_rows.append(row)

        # Sort professional/utox rows by client (column C), alphabetically
        professional_utox_rows.sort(key=lambda x: x[1].lower())

        # Reorder worksheet: sorted professional/utox first, then other rows
        new_row_pos = 2
        for original_row, _, _ in professional_utox_rows:
            for col in range(1, ws.max_column + 1):
                ws.cell(new_row_pos, col).value = row_data_map[original_row][col - 1]
            new_row_pos += 1

        for original_row in other_rows:
            for col in range(1, ws.max_column + 1):
                ws.cell(new_row_pos, col).value = row_data_map[original_row][col - 1]
            new_row_pos += 1

        # Assign staff: first 300 (sorted) to Rosanna, rest to Jasmine
        for idx in range(len(professional_utox_rows)):
            new_row_pos = idx + 2
            if idx < 300:
                split_assignment[new_row_pos] = "Rosanna"
            else:
                split_assignment[new_row_pos] = "Jasmine"
    
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

        # If split_professional_utox is enabled, use pre-computed assignment for those rows
        if split_professional_utox and row in split_assignment:
            staff = split_assignment[row]
            ws.cell(row, 1).value = staff
            continue

        # All IOP goes to Jasmine when split_professional_utox is enabled, except for Self-Pay IOP (CB)
        if split_professional_utox and "iop" in service_lower and group != "Self Pay":
            staff = "Jasmine"
            ws.cell(row, 1).value = staff
            continue

        # WM / OP WM Program Level → Melissa
        if 'program_level' in cols:
            pl_value = ws.cell(row, cols['program_level']).value
            if is_wm_program_level(pl_value):
                staff = "Melissa"

        if staff:
            ws.cell(row, 1).value = staff
            continue

        # Check if service is non-billable for this day using weekday rules
        is_non_billable = is_non_billable_service_for_weekday(service, day_of_week, php_on_monday)

        # Self-Pay (CB): Specific services are always billable every day regardless of weekday rules.
        if group == "Self Pay" and is_cb_billable_service(service):
            is_non_billable = False

        # When the Jasmine D/R/PHP option is active, Detox, Residential, and PHP
        # are billable on every weekday so they can be routed to Jasmine.
        if jasmine_detox_residential_php and _is_programming_service(service):
            is_non_billable = False

        # When "Give Jasmine IOP and all professional services" is active,
        # IOP, PHP, Acupuncture, and professional services are billable on every
        # weekday so they can be routed to Jasmine.
        if jasmine_iop_professional and (group == "Insurance" or group == ""):
            if ("iop" in service_lower or
                    "partial hospitalization" in service_lower or
                    "php" in service_lower or
                    service_lower.startswith("acupuncture") or
                    is_professional_service(service)):
                is_non_billable = False

        # When "Give utox to Jasmine" is active, drug screens are billable on
        # every weekday so they can be routed to Jasmine.
        if give_utox_to_jasmine and (group == "Insurance" or group == "") and is_drug_screen(service):
            is_non_billable = False

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

        # split_professional_utox: Professional and Utox services sorted by Client
        # First 300 to Rosanna, rest to Jasmine. All IOP goes to Jasmine.
        if not staff and split_professional_utox:
            if row in split_assignment:
                staff = split_assignment[row]
            elif "iop" in service_lower:
                staff = "Jasmine"

        # Melissa: (Detox or Residential) + (Aetna or Humana), but not drug screens
        if not staff:
            has_detox_res = ("detox" in service_lower or "residential" in service_lower)
            has_insurance = "aetna" in payer or "humana" in payer

            if has_detox_res and has_insurance and not is_drug_screen(service):
                staff = "Melissa"

        # CB: Self Pay
        if not staff and group == "Self Pay":
            staff = "CB"

        # Rosanna: Insurance + services based on active option
        if not staff and group == "Insurance":
            if rosanna_iop_php_acu:
                if ("iop" in service_lower or "partial hospitalization" in service_lower or
                        "php" in service_lower or service_lower.startswith("acupuncture")):
                    staff = "Rosanna"
            elif rosanna_php_iop_only:
                if ("iop" in service_lower or "partial hospitalization" in service_lower or
                        "php" in service_lower):
                    staff = "Rosanna"
            elif rosanna_iop_jasmine_php:
                if "iop" in service_lower:
                    staff = "Rosanna"
            elif jasmine_iop_professional:
                # IOP, PHP, Acupuncture, and all professional services go to Jasmine
                pass
            elif route_iop_acu_to_rosanna:
                if ("iop" in service_lower or
                        service_lower.startswith("acupuncture")):
                    staff = "Rosanna"
            elif jasmine_detox_residential_php:
                # PHP goes to Jasmine; IOP and Acupuncture still go to Rosanna
                if ("iop" in service_lower or
                        service_lower.startswith("acupuncture")):
                    staff = "Rosanna"
            else:
                if ("iop" in service_lower or
                        service_lower.startswith("acupuncture") or
                        "partial hospitalization" in service_lower):
                    staff = "Rosanna"

        # Jasmine: (Insurance or blank) + (Detox or Residential), but not drug screens
        # Also receives PHP when rosanna_iop_jasmine_php is enabled
        if not staff and (group == "Insurance" or group == ""):
            is_detox_res = ("detox" in service_lower or service_lower.startswith("residential"))
            is_php = ("partial hospitalization" in service_lower or "php" in service_lower)
            if ((is_detox_res or
                    (rosanna_iop_jasmine_php and is_php) or
                    (jasmine_detox_residential_php and is_php)) and
                    not is_drug_screen(service)):
                staff = "Jasmine"

        # Utox to Jasmine: drug screen rows go to Jasmine when checkbox is enabled
        if not staff and give_utox_to_jasmine and is_drug_screen(service):
            staff = "Jasmine"

        # Jasmine: inpatient (detox/residential) and all professional services
        if not staff and jasmine_inpatient_professional and (group == "Insurance" or group == ""):
            if ("detox" in service_lower or "residential" in service_lower or
                    is_professional_service(service)):
                staff = "Jasmine"

        # Jasmine: IOP, PHP, Acupuncture, and all professional services
        if not staff and jasmine_iop_professional and (group == "Insurance" or group == ""):
            if ("iop" in service_lower or
                    "partial hospitalization" in service_lower or
                    "php" in service_lower or
                    service_lower.startswith("acupuncture") or
                    is_professional_service(service)):
                staff = "Jasmine"

        # Fill remaining blanks
        if not staff:
            if jasmine_detox_residential_only or jasmine_detox_residential_php:
                staff = "Rosanna"
            else:
                staff = "Jasmine" if route_iop_acu_to_rosanna else "Rosanna"

        ws.cell(row, 1).value = staff

    print("Staff assignment complete")


def finalize_workbook(wb, exclude_drug_screens: bool = False, include_drug_screen_statuses: bool = True, include_batch_billings: bool = False):
    """Add Status/Comments columns and validation for Rosanna/Jasmine exports.

    Args:
        wb: Workbook to finalize.
        exclude_drug_screens: When True, omit 'Utox Batch' from dropdown (Rosanna only).
        include_drug_screen_statuses: When False, omit both 'Utox Batch' and 'Inclusive
            Services' from the dropdown. Used for Jasmine when give_utox_to_jasmine is False.
        include_batch_billings: When True, add 'Batch Billings' to the dropdown (Jasmine only).
    """
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
    ws_list = wb.create_sheet("Sheet2")
    status_items = ["Billed", "Unable to Bill", "Contractual Adj", "Incomplete Billings"]
    if include_drug_screen_statuses:
        if exclude_drug_screens:
            # Rosanna with drug screens excluded: keep Inclusive Services but drop Utox Batch
            status_items.append("Inclusive Services")
        else:
            # Rosanna (default) or Jasmine with utox: full dropdown
            status_items.append("Utox Batch")
            status_items.append("Inclusive Services")
    if include_batch_billings:
        status_items.append("Batch Billings")

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

def is_op_wm_program_level(cell_value) -> bool:
    """Return True if the Program Level cell contains 'OP WM'."""
    return "OP WM" in str(cell_value or "").upper()

def is_cb_billable_service(service: str) -> bool:
    """Return True if service is billable every day for CB (self pay).

    Billable CB services:
    - IOP
    - Group therapy
    - Individual therapy
    - Family therapy
    - Medication admin
    - MATS
    - Psych eval
    - Psych follow up
    """
    s = service.lower()
    cb_keywords = (
        "iop",
        "group",
        "individual",
        "family",
        "medication admin",
        "mats",
        "psych eval",
        "psych follow",
    )
    return any(keyword in s for keyword in cb_keywords)


def is_anthem_payer(payer: str) -> bool:
    """Return True if the payer contains 'anthem'."""
    return "anthem" in payer.lower()

def is_professional_service(service: str) -> bool:
    """Return True for professional outpatient services (not IOP/PHP/acupuncture/detox/residential/drug screen)."""
    s = service.lower()
    return (
        "iop" not in s and
        "partial hospitalization" not in s and
        "php" not in s and
        not s.startswith("acupuncture") and
        "detox" not in s and
        "residential" not in s and
        not is_drug_screen(service)
    )

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

def process_workbook(uploaded_file, exclude_drug_screens: bool = False,
                     exclude_optum: bool = False, give_utox_to_jasmine: bool = False,
                     route_iop_acu_to_rosanna: bool = False,
                     exclude_bcb_anthem_ct: bool = False,
                     rosanna_php_iop_only: bool = False,
                     jasmine_detox_residential_only: bool = False,
                     rosanna_iop_php_acu: bool = False,
                     jasmine_inpatient_professional: bool = False,
                     exclude_anthem_rosanna_jasmine_owm: bool = False,
                     php_on_monday: bool = False,
                     rosanna_iop_jasmine_php: bool = False,
                     jasmine_iop_professional: bool = False,
                     exclude_detox_residential: bool = False,
                     jasmine_detox_residential_php: bool = False,
                     split_professional_utox: bool = False):
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
        assign_staff(ws, date_token, give_utox_to_jasmine=give_utox_to_jasmine,
                     route_iop_acu_to_rosanna=route_iop_acu_to_rosanna,
                     rosanna_php_iop_only=rosanna_php_iop_only,
                     jasmine_detox_residential_only=jasmine_detox_residential_only,
                     rosanna_iop_php_acu=rosanna_iop_php_acu,
                     jasmine_inpatient_professional=jasmine_inpatient_professional,
                     php_on_monday=php_on_monday,
                     rosanna_iop_jasmine_php=rosanna_iop_jasmine_php,
                     jasmine_iop_professional=jasmine_iop_professional,
                     jasmine_detox_residential_php=jasmine_detox_residential_php,
                     split_professional_utox=split_professional_utox)

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

        # Rosanna gets her own report when any explicit Rosanna routing option
        # is checked; otherwise her caseload is redirected to Jasmine and her
        # report iteration finds no rows (and is skipped).
        rosanna_gets_report = (route_iop_acu_to_rosanna or rosanna_php_iop_only or
                               rosanna_iop_php_acu or rosanna_iop_jasmine_php or
                               split_professional_utox)
        staff_redirect = {} if rosanna_gets_report else STAFF_REDIRECT

        # Jasmine and CB always get individual reports; Rosanna gets one only
        # when activated above. Melissa, Unable to Bill, etc. stay Masters-only.
        for staff_name in ["Rosanna", "Jasmine", "CB"]:
            new_wb = openpyxl.Workbook()
            new_ws = new_wb.active
            new_ws.title = "Sheet1"

            for col in range(1, ws.max_column + 1):
                new_ws.cell(1, col).value = ws.cell(1, col).value

            new_row = 2
            for row in range(2, ws.max_row + 1):
                assigned_staff = ws.cell(row, 1).value
                effective_staff = staff_redirect.get(assigned_staff, assigned_staff)
                if effective_staff == staff_name:
                    # Exclude Optum utox rows from all individual workbooks
                    if (exclude_optum and service_col is not None and
                            payer_col is not None):
                        service_val = str(ws.cell(row, service_col).value or "")
                        payer_val = str(ws.cell(row, payer_col).value or "").lower()
                        if is_drug_screen(service_val) and "optum" in payer_val:
                            continue
                    if (exclude_drug_screens and assigned_staff == "Rosanna" and
                            service_col is not None):
                        service_val = str(ws.cell(row, service_col).value or "")
                        if is_drug_screen(service_val):
                            continue
                    if (exclude_bcb_anthem_ct and service_col is not None and
                            payer_col is not None):
                        service_val = str(ws.cell(row, service_col).value or "")
                        payer_val = str(ws.cell(row, payer_col).value or "")
                        if is_bcb_anthem_ct_php_res_detox(payer_val, service_val):
                            continue
                    if (rosanna_php_iop_only and assigned_staff == "Rosanna" and
                            service_col is not None):
                        service_val = str(ws.cell(row, service_col).value or "").lower()
                        if not ("iop" in service_val or "partial hospitalization" in service_val or "php" in service_val):
                            continue
                    if (rosanna_iop_php_acu and assigned_staff == "Rosanna" and
                            service_col is not None):
                        service_val = str(ws.cell(row, service_col).value or "").lower()
                        if not ("iop" in service_val or "partial hospitalization" in service_val or
                                "php" in service_val or service_val.startswith("acupuncture")):
                            continue
                    if (jasmine_detox_residential_only and assigned_staff == "Jasmine" and
                            service_col is not None):
                        service_val = str(ws.cell(row, service_col).value or "")
                        service_lower_val = service_val.lower()
                        if not ("detox" in service_lower_val or "residential" in service_lower_val):
                            if not (jasmine_inpatient_professional and
                                    ("detox" in service_lower_val or
                                     "residential" in service_lower_val or
                                     is_professional_service(service_val))):
                                continue
                    # Jasmine's report is limited to Detox/Residential/PHP. This
                    # matches on staff_name (not assigned_staff) so that rows
                    # redirected into Jasmine from Rosanna are dropped too.
                    if (jasmine_detox_residential_php and staff_name == "Jasmine" and
                            service_col is not None):
                        service_val = str(ws.cell(row, service_col).value or "").lower()
                        if not ("detox" in service_val or "residential" in service_val or
                                "partial hospitalization" in service_val or "php" in service_val):
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
                    # Overwrite the staff column so redirected rows show the
                    # owning staff name (e.g. Rosanna rows folded into Jasmine
                    # should read "Jasmine" in Jasmine's individual report).
                    new_ws.cell(new_row, 1).value = staff_name
                    new_row += 1

            if new_row == 2:
                continue

            if staff_name == "Rosanna":
                finalize_workbook(new_wb, exclude_drug_screens=exclude_drug_screens)
            elif staff_name == "Jasmine":
                # Jasmine only gets "Utox Batch" / "Inclusive Services" when
                # the "Give utox to Jasmine" checkbox is checked.
                finalize_workbook(new_wb, include_drug_screen_statuses=give_utox_to_jasmine,
                                  include_batch_billings=True)
            elif staff_name == "CB":
                finalize_workbook(new_wb)

            output = io.BytesIO()
            new_wb.save(output)
            output.seek(0)
            output_filename = f"{filename_prefix}{staff_name}.xlsx"
            output_files[output_filename] = output

        # Mirror the redirect in the Masters report so it matches the individual
        # reports. When Rosanna has her own report this is a no-op.
        for row in range(2, ws.max_row + 1):
            assigned_staff = ws.cell(row, 1).value
            if assigned_staff in staff_redirect:
                ws.cell(row, 1).value = staff_redirect[assigned_staff]

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

exclude_optum = st.checkbox(
    "Exclude Optum insurance",
    value=False,
    help="When checked, utox (drug screen) rows with Optum as the payer will be excluded from all individual staff workbooks."
)

give_utox_to_jasmine = st.checkbox(
    "Give utox to Jasmine",
    value=False,
    help="When checked, utox (drug screen) rows will be assigned to Jasmine in the master spreadsheet and included in Jasmine's workbook."
)

route_iop_acu_to_rosanna = st.checkbox(
    "Route IOP and Acupuncture to Rosanna (all other services to Jasmine)",
    value=False,
    help="When checked, only IOP and Acupuncture services go to Rosanna. All remaining services (including Partial Hospitalization) default to Jasmine instead of Rosanna."
)

exclude_bcb_anthem_ct = st.checkbox(
    "Exclude BCB Anthem CT for PHP, Residential, and Detox",
    value=False,
    help="When checked, rows where the payer is BCB Anthem CT and the service is PHP (Partial Hospitalization), Residential, or Detox will be excluded from all individual staff workbooks."
)

rosanna_php_iop_only = st.checkbox(
    "Give Rosanna only PHP and IOP",
    value=False,
    help="When checked, Rosanna only receives PHP (Partial Hospitalization) and IOP services. Acupuncture will not be routed to Rosanna."
)

jasmine_detox_residential_only = st.checkbox(
    "Give Jasmine only Detox and Residential",
    value=False,
    help="When checked, Jasmine only receives Detox and Residential services. She will not be used as the fallback default for unmatched services."
)

rosanna_iop_php_acu = st.checkbox(
    "Give Rosanna IOP, PHP, and Acupuncture",
    value=False,
    help="When checked, Rosanna receives IOP, PHP (Partial Hospitalization), and Acupuncture services. Her workbook is filtered to only these three service types."
)

jasmine_inpatient_professional = st.checkbox(
    "Give Jasmine Inpatient (Detox/Residential) and all Professional services",
    value=False,
    help="When checked, Jasmine receives inpatient services (Detox and Residential) as well as all professional outpatient services. Also expands her workbook when 'Give Jasmine only Detox and Residential' is active."
)

exclude_anthem_rosanna_jasmine_owm = st.checkbox(
    "Remove Anthem from Rosanna and Jasmine reports",
    value=False,
    help="When checked, rows where the payer contains 'Anthem' will be excluded from Rosanna's and Jasmine's workbooks. Anthem rows are still retained in the Masters report."
)

php_on_monday = st.checkbox(
    "Run PHP (Partial Hospitalization) on Mondays",
    value=False,
    help="When checked, Partial Hospitalization services will be treated as billable on Mondays instead of being marked 'Unable to Bill'."
)

rosanna_iop_jasmine_php = st.checkbox(
    "Give Rosanna IOP and Jasmine PHP",
    value=False,
    help="When checked, IOP services go to Rosanna and PHP (Partial Hospitalization) services go to Jasmine. All day-of-week rules still apply."
)

jasmine_iop_professional = st.checkbox(
    "Give Jasmine IOP and all Professional services",
    value=False,
    help="When checked, IOP, PHP (Partial Hospitalization), Acupuncture, and all professional outpatient services are routed to Jasmine. Rosanna does not receive IOP, PHP, or Acupuncture under this option."
)

exclude_detox_residential = st.checkbox(
    "Don't give anyone Detox or Residential",
    value=False,
    help="When checked, Detox and Residential service rows are excluded from all individual staff workbooks. They still appear in the Masters report."
)

jasmine_detox_residential_php = st.checkbox(
    "Give Jasmine only Detox, Residential, and PHP",
    value=False,
    help="When checked, Detox, Residential, and PHP (Partial Hospitalization) are treated as billable on every weekday and routed to Jasmine. Jasmine's report is limited to just these three services. Rosanna does not receive a separate report unless one of her own options is checked; all other services stay in the Masters report only."
)

split_professional_utox = st.checkbox(
    "Give Rosanna first 300 Professional/Utox (sorted by Client), Jasmine gets rest + all IOP",
    value=False,
    help="When checked, all Professional and Utox services are sorted by Client (column C) A-Z. First 300 rows go to Rosanna, remaining go to Jasmine. All IOP goes to Jasmine. Creates separate reports for both."
)

if uploaded_file is not None:
    try:
        validate_uploaded_file(uploaded_file)
        st.info("Processing your file...")
        output_files, invalid_count, date_token, wm_count = process_workbook(
            uploaded_file,
            exclude_drug_screens=exclude_drug_screens,
            exclude_optum=exclude_optum,
            give_utox_to_jasmine=give_utox_to_jasmine,
            route_iop_acu_to_rosanna=route_iop_acu_to_rosanna,
            exclude_bcb_anthem_ct=exclude_bcb_anthem_ct,
            rosanna_php_iop_only=rosanna_php_iop_only,
            jasmine_detox_residential_only=jasmine_detox_residential_only,
            rosanna_iop_php_acu=rosanna_iop_php_acu,
            jasmine_inpatient_professional=jasmine_inpatient_professional,
            exclude_anthem_rosanna_jasmine_owm=exclude_anthem_rosanna_jasmine_owm,
            php_on_monday=php_on_monday,
            rosanna_iop_jasmine_php=rosanna_iop_jasmine_php,
            jasmine_iop_professional=jasmine_iop_professional,
            exclude_detox_residential=exclude_detox_residential,
            jasmine_detox_residential_php=jasmine_detox_residential_php,
            split_professional_utox=split_professional_utox,
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
                logger.info(f"User '{st.session_state.get('username', 'unknown')}' downloaded: {output_filename}")
    
    except Exception as e:
        logger.error(f"Error processing file for user '{st.session_state.get('username', 'unknown')}': {str(e)}")
        st.error(f"Error: {str(e)}")
