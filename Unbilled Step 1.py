import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from pathlib import Path
import re
from datetime import datetime

def extract_date_from_filename(filename):
    """Extract MMDDYYYY date from filename"""
    match = re.search(r'(\d{8})', filename)
    if match:
        return match.group(1)
    return None

def get_save_folder(workbook_path, date_token):
    """Build save path: ...\\Unbilled Reports\\YYYY\\MM - Month YYYY\\MMDDYYYY\\"""
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
        service = str(ws.cell(row, cols['service']).value or "")
        payer = str(ws.cell(row, cols['payer']).value or "")
        billing_provider = str(ws.cell(row, cols['billing_provider']).value or "").strip()
        
        staff = None
        
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
        
        # Rosanna_1: Insurance + (IOP or Acupuncture or Partial Hospitalization)
        if not staff and group == "Insurance":
            if ("IOP" in service or 
                service.startswith("Acupuncture") or 
                "Partial Hospitalization" in service):
                staff = "Rosanna"
        
        # Melissa: (Detox or Residential but NOT Drug Screen) + (Aetna or Humana)
        # CHECK MELISSA BEFORE JASMINE - she's more specific
        if not staff:
            has_detox_res = ("Detox" in service or "Residential" in service)
            no_drug_screen = "Drug Screen" not in service
            has_insurance = "Aetna" in payer or "Humana" in payer
            
            if has_detox_res and no_drug_screen and has_insurance:
                staff = "Melissa"
        
        # Jasmine: (Insurance or blank) + (Detox or Drug Screen 13 Panel or Residential)
        if not staff and (group == "Insurance" or group == ""):
            if (service.startswith("Detox") or 
                service.startswith("Drug Screen 13 Panel") or 
                service.startswith("Residential")):
                staff = "Jasmine"
        
        # Rosanna_2: Fill remaining blanks
        if not staff:
            staff = "Rosanna"
        
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
    
    # Add data validation to Status column
    dv = DataValidation(type="list", formula1="=Sheet2!$A$1:$A$4", allow_blank=True)
    ws.add_data_validation(dv)
    
    last_row = ws.max_row
    dv.add(f'E2:E{last_row}')
    
    # Auto-fit columns
    for col in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col)].auto_size = True

def export_staff_workbooks(wb, wb_path, date_token):
    """Export separate workbooks for Rosanna, Jasmine, CB"""
    
    ws = wb.active
    save_folder = get_save_folder(wb_path, date_token)
    
    for staff_name in ["Rosanna", "Jasmine", "CB"]:
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
            if ws.cell(row, 1).value == staff_name:
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
        raise ValueError("Could not extract date (MMDDYYYY) from filename")
    
    print(f"Processing file with date: {date_token}")
    
    # Step 1: Extract invalid
    step_1_extract_invalid(ws)
    
    # Step 2-6: Assign staff
    assign_staff(ws, date_token)
    
    # Save main workbook
    wb.save(workbook_path)
    print(f"Saved main workbook: {workbook_path}")
    
    # Step 7: Export individual workbooks
    export_staff_workbooks(wb, workbook_path, date_token)
    
    print("\n✓ All done!")

if __name__ == "__main__":
    # UPDATE THIS LINE with your actual file path
    workbook_path = r"C:\Users\phogo\OneDrive\Billing App\Unbilled Reports\2026\01 - January 2026\01252026\Unbilled Revenue By Resident and Funding Type 01252026 - Masters.xlsx"
    
    main(workbook_path)