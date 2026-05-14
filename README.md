# Billing App

An automated billing processor for unbilled revenue reports.

## Features

- Processes Excel workbooks with billing data
- Extracts invalid billing records to a separate sheet
- Assigns staff members based on business rules
- Applies weekday-based non-billable service logic using date from filename
- Generates separate workbooks for each staff member

## Weekday-Based Non-Billable Rules

The system uses the date extracted from the filename (MMDDYYYY format) to determine weekday-based billing rules.

Services fall into two categories:

- **Programming**: Detox, Residential, Partial Hospitalization (PHP)
- **Professional**: all other services (including IOP and Acupuncture)

Weekly billing schedule:

- **Monday**: Professional only (Programming non-billable)
- **Tuesday**: Programming only (Professional non-billable)
- **Wednesday**: Professional only (Programming non-billable)
- **Thursday**: Programming + Professional
- **Friday**: Programming + Professional
- **Saturday/Sunday**: All services billed except e-care

**E-care**: billable on Tuesdays only. The system recognizes 'e-care', 'e care', 'ecare', and 'extended care' (case-insensitive).

The "Run PHP on Mondays" option exempts Partial Hospitalization from the Monday restriction.

## Reports

Individual workbooks are generated for **Jasmine** and **CB** only. All other staff
(Melissa, Rosanna, etc.) are still assigned in the Masters workbook but do not
receive separate reports.

### Fallback Behavior

If the filename does not contain a valid 8-digit date (MMDDYYYY), the system will:
- Use today's date for determining weekday rules
- Show a warning in the UI (for Streamlit app)
- Print a warning to console (for command-line script)

## Usage

### Streamlit Web App

```bash
streamlit run app.py
```

Upload an Excel file with MMDDYYYY in the filename (e.g., `Report_01252026.xlsx`).

### Command-Line Script

```bash
python "Unbilled Step 1.py"
```

Update the `workbook_path` variable at the bottom of the script with your file path.

## Testing

Run the unit tests for weekday rules:

```bash
python tests/test_weekday_rules.py
```

Or use pytest if available:

```bash
pytest tests/test_weekday_rules.py -v
```

## Requirements

```
streamlit
openpyxl
```

Install dependencies:

```bash
pip install -r requirements.txt
```

## File Structure

- `app.py` - Streamlit web application
- `Unbilled Step 1.py` - Command-line processing script
- `billing_rules.py` - Business logic for weekday-based non-billable determination
- `tests/test_weekday_rules.py` - Unit tests for billing rules
- `requirements.txt` - Python dependencies

## License

Proprietary
