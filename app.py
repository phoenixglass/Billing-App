import streamlit as st
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import re
from datetime import datetime
import io
import tempfile

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

def assign_staff(ws):
    """Assign staff names based on business rules"""
    
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
    
    # Get today's day of week (0=Monday, 6=Sunday)
    today = datetime.now()
    day_of_week = today.weekday()  # 0=Monday, 1=Tuesday, ..., 4=Friday
    
    # Define non-billable services by day (case-insensitive matching)
    non_billable_services = {
        0: ["detox", "residential", "partial hospitalization", "e-care"],  # Monday
        1: [],  # Tuesday - bill all
        2: ["e-care"],  # Wednesday
        3: ["e-care"],  # Thursday
        4: ["e-care"],  # Friday
        5: ["e-care"],  # Saturday
        6: ["e-care"],  # Sunday
    }
    
    non_billable = non_billable_services.get(day_of_week, [])
    
    # Assign staff
    for row in range(2, ws.max_row + 1):
        group = str(ws.cell(row, cols['group']).value or "").strip()
        service = str(ws.cell(row, cols['service']).value or "").lower()
        payer = str(ws.cell(row, cols['payer']).value or "").lower()
        
        staff = None
        
        # Check if service is non-billable for this day (case-insensitive)
        is_non_billable = any(non_bill_service in service for non_bill_service in non_billable)
        
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
            has_detox_res = ("detox" in service or "residential" in service)
            has_insurance = "aetna" in payer or "humana" in payer
            
            if has_detox_res and has_insurance:
                staff = "Melissa"
        
        # CB: Self Pay
        if not staff and group == "Self Pay":
            staff = "CB"
        
        # Rosanna_1: Insurance + (IOP or Acupuncture or Partial Hospitalization)
        if not staff and group == "Insurance":
            if ("iop" in service or 
                service.startswith("acupuncture") or 
                "partial hospitalization" in service):
                staff = "Rosanna"
        
        # Jasmine: (Insurance or blank) + (Detox or Drug Screen 13 Panel or Residential)
        if not staff and (group == "Insurance" or group == ""):
            if ("detox" in service or 
                service.startswith("drug screen 13 panel") or 
                service.startswith("residential")):
                staff = "Jasmine"
        
        # Rosanna_2: Fill remaining blanks
        if not staff:
            staff = "Rosanna"
        
        ws.cell(row, 1).value = staff
    
    print("Staff assignment complete")

def finalize_workbook(wb):
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
    
    # Widen columns to fit all text
    for col in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col)].auto_size = True

def process_workbook(uploaded_file):
    """Process the uploaded workbook"""
    
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
    assign_staff(ws)
    
    output_files = {}
    
    for staff_name in ["Rosanna", "Jasmine", "CB", "Melissa", "Unable to Bill"]:
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active
        new_ws.title = "Sheet1"
        
        for col in range(1, ws.max_column + 1):
            new_ws.cell(1, col).value = ws.cell(1, col).value
        
        new_row = 2
        for row in range(2, ws.max_row + 1):
            if ws.cell(row, 1).value == staff_name:
                for col in range(1, ws.max_column + 1):
                    new_ws.cell(new_row, col).value = ws.cell(row, col).value
                new_row += 1
        
        if new_row == 2:
            continue
        
        if staff_name in ["Rosanna", "Jasmine"]:
            finalize_workbook(new_wb)
        
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
    
    return output_files, invalid_count, date_token

# UI
st.markdown("### Upload your billing file")
uploaded_file = st.file_uploader(
    "Choose an Excel file (must have MMDDYYYY in filename)",
    type="xlsx"
)

if uploaded_file is not None:
    try:
        st.info("Processing your file...")
        output_files, invalid_count, date_token = process_workbook(uploaded_file)
        
        st.success("✓ Processing complete!")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Date Extracted", date_token)
        with col2:
            st.metric("Invalid Rows", invalid_count or 0)
        with col3:
            st.metric("Files Generated", len(output_files))
        
        st.markdown("---")
        st.markdown("### Download Results")
        
        for output_filename, file_bytes in output_files.items():
            st.download_button(
                label=f"Download {output_filename}",
                data=file_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    
    except Exception as e:
        st.error(f"Error: {str(e)}")