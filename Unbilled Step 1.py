import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from pathlib import Path
import re
from datetime import datetime
from billing_rules import is_non_billable_service_for_weekday

stubs_required = False  # no Streamlit in the standalone script; keep for parity if needed

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
    """Check if service is a Programming service: Detox, Residential, or PHP."""
    s = service.lower()
    return (
        'detox' in s
        or 'residential' in s
        or 'partial hospitalization' in s
        or 'php' in s
    )

def is_non_billable_service_for_weekday(
    service: str, weekday: int, php_on_monday: bool = False
) -> bool:
    """
    Determine if a service is non-billable for a given weekday.

    Weekly billing schedule:
    - Monday (0):    Professional only        (Programming non-billable)
    - Tuesday (1):   Programming only         (Professional non-billable)
    - Wednesday (2): Professional only        (Programming non-billable)
    - Thursday (3):  Programming + Professional
    - Friday (4):    Programming + Professional
    - Saturday/Sunday (5,6): all services billable except e-care

    Service categories:
    - Programming: Detox, Residential, Partial Hospitalization (PHP)
    - Professional: all other services (including IOP and Acupuncture)

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

    is_programming = _is_programming_service(service_lower)

    # Monday and Wednesday: Professional only
    if weekday in (0, 2):
        if not is_programming:
            return False
        if weekday == 0 and php_on_monday and (
            'partial hospitalization' in service_lower or 'php' in service_lower
        ):
            return False
        return True

    # Tuesday: Programming only
    if weekday == 1:
        return not is_programming

    # Thursday and Friday: Programming + Professional both billable
    if weekday in (3, 4):
        return False

    # Saturday and Sunday: everything billable (e-care handled above)
    return False

def is_wm_program_level(cell_value) -> bool:
    """Return True if the Program Level cell contains 'WM'."""
    return "WM" in str(cell_value or "").upper()

def get_save_folder(workbook_path, date_token):
    """Build save path: ...\\Unbilled Reports\\YYYY\\MM - Month YYYY\\MMDDYYYY"""
    wb_path = Path(workbook_path)

    # Find "Unbilled Reports" in the path
    parts = wb_path.parts
    try:
        idx = [i for i, p in enumerate(parts) if 'Unbilled Reports' in p][0]
        base = Path(*parts[:idx+1])
    except IndexError:
        raise ValueError("Workbook must be saved in a folder containing 'Unbilled Reports'")

    # Parse date token: MMDDYYYY
    month = date_token[:2]
    day = date_token[2:4]
    year = date_token[4:]

    dt = datetime.strptime(date_token, '%m%d%Y')
    month_name = dt.strftime('%B')

    # Build: YYYY\\MM - Month YYYY\\MMDDYYYY
    save_path = base / year / f"{month} - {month_name} {year}" / date_token
    save_path.mkdir(parents=True, exist_ok=True)

    return save_path

def step_1_extract_invalid(ws):
    """Insert Staff/Status, extract Invalid rows to new sheet, delete from main"""

    # Insert Staff/Status column if needed
    if ws.cell(1, 1).value != "Staff/Status":
        ws.insert_cols(1)
        ws.cell(1, 1).value = "Staff/Status"

    # Find Billable Status column
    billable_col = None
    for col in range(1, ws.max_column + 1):
        if ws.cell(1, col).value == "Billable Status":
            billable_col = col
            break

    if not billable_col:
        raise ValueError("Billable Status column not found")

    # Find all invalid rows
    invalid_rows = []
    for row in range(2, ws.max_row + 1):
        if ws.cell(row, billable_col).value == "Invalid for Billing":
            invalid_rows.append(row)

    if not invalid_rows:
        print("No invalid rows found")
        return None

    # Create Invalid sheet
    wb = ws.parent
    if "Invalid" in wb.sheetnames:
        del wb["Invalid"]

    ws_invalid = wb.create_sheet("Invalid")

    # Copy header
    for col in range(1, ws.max_column + 1):
        ws_invalid.cell(1, col).value = ws.cell(1, col).value

    # Copy invalid rows
    for idx, row_num in enumerate(invalid_rows, start=2):
        for col in range(1, ws.max_column + 1):
            ws_invalid.cell(idx, col).value = ws.cell(row_num, col).value

    # Delete invalid rows (backwards to preserve row numbers)
    for row_num in reversed(invalid_rows):
        ws.delete_rows(row_num)

    # Format headers
    ws.cell(1, 1).font = openpyxl.styles.Font(bold=True)
    ws.auto_filter.ref = ws.dimensions

    print(f"Extracted {len(invalid_rows)} invalid rows to 'Invalid' sheet")
    return len(invalid_rows)

def _is_ecare(service: str) -> bool:
    """Check if service is e-care (handles common variants)"""
    # Handle common e-care variants: "e-care", "e care", "ecare", "extended care"
    s = service.lower()
    return any(variant in s for variant in ["e-care", "e care", "ecare", "extended care"])

def is_non_billable_service_for_weekday(
    service: str, weekday: int, php_on_monday: bool = False
) -> bool:
    """
    Determine if a service is non-billable for a given weekday.

    Weekly billing schedule:
    - Monday (0):    Professional only        (Programming non-billable)
    - Tuesday (1):   Programming only         (Professional non-billable)
    - Wednesday (2): Professional only        (Programming non-billable)
    - Thursday (3):  Programming + Professional
    - Friday (4):    Programming + Professional
    - Saturday/Sunday (5,6): all services billable except e-care

    Service categories:
    - Programming: Detox, Residential, Partial Hospitalization (PHP)
    - Professional: all other services (including IOP and Acupuncture)

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

    is_programming = _is_programming_service(service_lower)

    # Monday and Wednesday: Professional only
    if weekday in (0, 2):
        if not is_programming:
            return False
        if weekday == 0 and php_on_monday and (
            "partial hospitalization" in service_lower or "php" in service_lower
        ):
            return False
        return True

    # Tuesday: Programming only
    if weekday == 1:
        return not is_programming

    # Thursday and Friday: Programming + Professional both billable
    if weekday in (3, 4):
        return False

    # Saturday and Sunday: everything billable (e-care handled above)
    return False

# Rosanna's caseload is currently redirected to Jasmine. The Rosanna routing
# options are kept intact so they can be re-enabled for future exceptions —
# clear this mapping to restore Rosanna's own assignments and report.
STAFF_REDIRECT = {"Rosanna": "Jasmine"}


def assign_staff(ws, date_token: str = None, route_iop_acu_to_rosanna: bool = False,
                 rosanna_php_iop_only: bool = False,
                 jasmine_detox_residential_only: bool = False):
    """Assign staff names based on business rules

    Args:
        ws: Worksheet to process
        date_token: Date string in MMDDYYYY format (from filename). If None, uses current date.
        route_iop_acu_to_rosanna: If True, only IOP and Acupuncture go to Rosanna and all
            remaining services default to Jasmine instead of Rosanna.
        rosanna_php_iop_only: If True, only PHP (Partial Hospitalization) and IOP go to
            Rosanna (Acupuncture excluded).
        jasmine_detox_residential_only: If True, Jasmine only receives Detox and Residential
            rows; she is not used as the fallback default for unmatched services.
    """
    
    # Find column indices (after Staff/Status insert, columns shift by 1)
    cols = {}
    for col in range(1, ws.max_column + 1):
        header = ws.cell(1, col).value
        if header == "GROUPFLD2":
            cols['group'] = col
        elif header == "GROUPFLD1":
            cols['groupfld1'] = col
        elif header == "Service":
            cols['service'] = col
        elif header == "Payer":
            cols['payer'] = col
        elif header == "Billing Provider":
            cols['billing_provider'] = col
        elif header == "Program Level":
            cols['program_level'] = col

    if not all(k in cols for k in ['group', 'groupfld1', 'service', 'payer', 'billing_provider']):
        raise ValueError("Missing required columns: GROUPFLD2, GROUPFLD1, Service, Payer, or Billing Provider")
    
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
        groupfld1 = str(ws.cell(row, cols['groupfld1']).value or "").strip()
        service = str(ws.cell(row, cols['service']).value or "").strip()
        service_lower = service.lower()
        payer = str(ws.cell(row, cols['payer']).value or "")
        payer_lower = payer.lower()
        billing_provider = str(ws.cell(row, cols['billing_provider']).value or "").strip()

        staff = None

        # WM Program Level: always assigned to Melissa
        if 'program_level' in cols:
            if is_wm_program_level(ws.cell(row, cols['program_level']).value):
                staff = "Melissa"

        if staff:
            ws.cell(row, 1).value = staff
            continue

        # Check if service is non-billable for this day using weekday rules
        is_non_billable = is_non_billable_service_for_weekday(service, day_of_week)

        # Unable to Bill: O'Flynn + OP Chappaqua or OP NYC
        if billing_provider == "O'Flynn, Karen":
            if groupfld1 == "OP Chappaqua" or groupfld1 == "OP NYC":
                staff = "Unable to Bill"
        
        # Unable to Bill: Non-billable service for this day of week
        if not staff and is_non_billable:
            staff = "Unable to Bill"
        
        # CB: Self Pay
        if not staff and group == "Self Pay":
            staff = "CB"

        # Rosanna: Insurance + services based on active option
        if not staff and group == "Insurance":
            if rosanna_php_iop_only:
                if ("iop" in service_lower or "partial hospitalization" in service_lower):
                    staff = "Rosanna"
            elif route_iop_acu_to_rosanna:
                if ("iop" in service_lower or
                    service_lower.startswith("acupuncture")):
                    staff = "Rosanna"
            else:
                if ("iop" in service_lower or
                    service_lower.startswith("acupuncture") or
                    "partial hospitalization" in service_lower):
                    staff = "Rosanna"

        # Melissa: (Detox or Residential but NOT Drug Screen) + (Aetna or Humana)
        # CHECK MELISSA BEFORE JASMINE - she's more specific
        if not staff:
            has_detox_res = ("detox" in service_lower or "residential" in service_lower)
            no_drug_screen = "drug screen" not in service_lower
            has_insurance = "aetna" in payer_lower or "humana" in payer_lower

            if has_detox_res and no_drug_screen and has_insurance:
                staff = "Melissa"

        # Jasmine: (Insurance or blank) + (Detox or Residential)
        if not staff and (group == "Insurance" or group == ""):
            if (service_lower.startswith("detox") or
                service_lower.startswith("residential")):
                staff = "Jasmine"

        # Fill remaining blanks
        if not staff:
            if jasmine_detox_residential_only:
                staff = "Rosanna"
            else:
                staff = "Jasmine" if route_iop_acu_to_rosanna else "Rosanna"

        ws.cell(row, 1).value = staff

    print("Staff assignment complete")

def finalize_workbook(wb):
    """Add Status/Comments columns and validation for Rosanna/Jasmine exports"""
    ws = wb.active

    # Insert two columns at E
    ws.insert_cols(5, 2)
    ws.cell(1, 5).value = "Status"
    ws.cell(1, 6).value = "Comments"

    # Create Sheet2 with validation list
    ws_list = wb.create_sheet("Sheet2")
    ws_list['A1'] = "Billed"
    ws_list['A2'] = "Unable to Bill"
    ws_list['A3'] = "Contractual Adj"
    ws_list['A4'] = "Incomplete Billings"
    ws_list['A5'] = "Utox Batch"
    ws_list['A6'] = "Inclusive Services"

    # Add data validation to Status column
    dv = DataValidation(type="list", formula1="=Sheet2!$A$1:$A$6", allow_blank=True)
    ws.add_data_validation(dv)

    last_row = ws.max_row
    dv.add(f'E2:E{last_row}')

    # Auto-fit columns
    for col in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col)].auto_size = True

def export_staff_workbooks(wb, wb_path, date_token):
    """Export separate workbooks for Jasmine and CB only.

    All staff (Melissa, Rosanna, Unable to Bill, etc.) are still assigned in the
    Masters workbook, but only Jasmine and the CB team receive individual reports.
    """

    ws = wb.active
    save_folder = get_save_folder(wb_path, date_token)

    # Find Program Level column for WM filtering
    program_level_col = None
    for col in range(1, ws.max_column + 1):
        if ws.cell(1, col).value == "Program Level":
            program_level_col = col
            break

    # Count and warn about WM rows
    wm_count = 0
    if program_level_col is not None:
        for row in range(2, ws.max_row + 1):
            if is_wm_program_level(ws.cell(row, program_level_col).value):
                wm_count += 1
    if wm_count > 0:
        print(
            f"WARNING: {wm_count} row(s) with 'WM' in the Program Level column were found. "
            "These rows are assigned to Melissa and appear only in the Masters report. "
            "Only Melissa is authorized to bill WM."
        )

    for staff_name in ["Jasmine", "CB"]:
        # Create new workbook
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active
        new_ws.title = "Sheet1"

        # Copy header
        for col in range(1, ws.max_column + 1):
            new_ws.cell(1, col).value = ws.cell(1, col).value

        # Copy matching rows
        new_row = 2
        for row in range(2, ws.max_row + 1):
            assigned_staff = ws.cell(row, 1).value
            effective_staff = STAFF_REDIRECT.get(assigned_staff, assigned_staff)
            if effective_staff == staff_name:
                # Skip WM program level rows for all staff except Melissa
                # (Masters retains all rows; only Melissa bills WM)
                if (staff_name != "Melissa" and
                        program_level_col is not None and
                        is_wm_program_level(ws.cell(row, program_level_col).value)):
                    continue
                for col in range(1, ws.max_column + 1):
                    new_ws.cell(new_row, col).value = ws.cell(row, col).value
                new_row += 1

        if new_row == 2:  # No rows found
            print(f"No rows for {staff_name}, skipping")
            continue

        # Finalize for Rosanna and Jasmine only
        if staff_name in ["Rosanna", "Jasmine"]:
            finalize_workbook(new_wb)

        # Save
        save_path = save_folder / f"Unbilled Revenue By Resident and Funding Type {date_token} - {staff_name}.xlsx"
        new_wb.save(save_path)
        print(f"Saved {save_path}")

def main(workbook_path):
    """Main workflow"""

    # Load workbook
    wb = openpyxl.load_workbook(workbook_path)

    # Find the data sheet (first sheet or by name pattern)
    ws = wb.active

    # Extract date token
    date_token = extract_date_from_filename(Path(workbook_path).name)
    if not date_token:
        print("Warning: Could not extract date (MMDDYYYY) from filename, will use today's date")
    
    print(f"Processing file with date: {date_token if date_token else 'today'}")
    
    # Step 1: Extract invalid
    step_1_extract_invalid(ws)

    # Step 2-6: Assign staff
    assign_staff(ws, date_token)

    # Step 7: Export individual workbooks (only if date_token is available)
    # Runs before the redirect so the Rosanna routing options stay effective.
    if date_token:
        export_staff_workbooks(wb, workbook_path, date_token)
    else:
        print("Skipping individual workbook export due to missing date token")

    # Redirect Rosanna's caseload into Jasmine in the Masters report so it
    # matches the individual reports. The Rosanna routing options stay intact
    # for future exceptions.
    for row in range(2, ws.max_row + 1):
        assigned_staff = ws.cell(row, 1).value
        if assigned_staff in STAFF_REDIRECT:
            ws.cell(row, 1).value = STAFF_REDIRECT[assigned_staff]

    # Save main workbook
    wb.save(workbook_path)
    print(f"Saved main workbook: {workbook_path}")
    
    print("\n✓ All done!")

if __name__ == "__main__":
    # Optionally allow running from CLI
    import sys
    if len(sys.argv) > 1:
        main(sys.argv[1])
