"""
Small helper to create sample Excel files to test Monday vs Tuesday rules.

Run: python tests/create_sample_workbooks.py
Requires: openpyxl
"""
from openpyxl import Workbook

def create_sample(filename, service_rows):
    wb = Workbook()
    ws = wb.active
    headers = ["Staff/Status", "GROUPFLD2", "GROUPFLD1", "Service", "Payer", "Billing Provider"]
    ws.append(headers)
    for row in service_rows:
        # fill other fields with simple values where needed
        ws.append(["", row.get("group", ""), row.get("groupfld1", ""), row["service"], row.get("payer", ""), row.get("billing_provider", "")])
    wb.save(filename)
    print(f"Created {filename}")

if __name__ == "__main__":
    monday_services = [
        {"service": "Partial Hospitalization - Program", "group": "Insurance", "payer": "Aetna"},
        {"service": "Residential Program", "group": "Insurance", "payer": "Humana"},
        {"service": "Detox Admission", "group": "Insurance", "payer": "Aetna"},
        {"service": "E-Care Individual", "group": "Insurance", "payer": "Aetna"},
        {"service": "Regular Therapy", "group": "Insurance", "payer": "Aetna"},
    ]
    tuesday_services = monday_services[:]  # same rows for comparison
    create_sample("sample_monday_01012024.xlsx", monday_services)
    create_sample("sample_tuesday_01022024.xlsx", tuesday_services)
