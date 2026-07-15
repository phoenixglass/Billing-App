# Billing App

An automated billing processor for unbilled revenue reports.

## Features

- Processes Excel workbooks with billing data
- Extracts invalid billing records to a separate sheet
- Assigns staff members based on business rules
- Applies weekday-based non-billable service logic using date from filename
- Generates separate workbooks for each staff member

## Daily Billing Rules

The system uses the date extracted from the filename (MMDDYYYY format) to determine
the day's schedule automatically — there are no manual day-of-week toggles.

**Self Pay**: every row with `GROUPFLD2` = "Self Pay" goes to **CB**. Every Self Pay
service bills every day of the week, with no exceptions (including e-care).

**Insurance** (`GROUPFLD2` = "Insurance") rows are split between **Rosanna**,
**Joshua**, and **Jasmine**. `GROUPFLD2` values other than "Insurance" or "Self
Pay" never reach Rosanna, Joshua, or Jasmine — they are marked Unable to Bill.

- **Professional** services (identified by the `Claim Type` column equal to
  `CMS-1500`) bill every day of the week. The Insurance + CMS-1500 rows ("the
  professional pool") are sorted alphabetically by `Client`. On Monday, Rosanna
  receives the first N rows of that sorted pool per the weekday cap below; on
  Tuesday, Thursday, and Friday, Joshua receives the first N rows instead. The
  rest of the pool (and all of it on Wednesday/weekends) goes to Jasmine.
- **Programming** services (Detox, Residential) bill Tuesday, Thursday,
  Friday, and weekends; they are Unable to Bill on Monday and Wednesday. All
  billable Programming rows go to Jasmine.
- **IOP** (including Telemed IOP) bills every day of the week, with no
  exceptions, and always goes to Jasmine, bypassing the professional
  pool/Rosanna/Joshua split even if Claim Type is CMS-1500.
- **E-care** bills on Tuesdays only (regardless of Claim Type). Billable e-care
  rows go to Jasmine.
- **PHP** (Partial Hospitalization) always goes to **Melissa**, every day — see
  the Melissa section below. It is not part of the Programming bucket above and
  never reaches Rosanna, Joshua, or Jasmine.

Rosanna's and Joshua's professional-pool cap by weekday:

| Day       | Capped staff | Cap      | Report                    |
|-----------|--------------|----------|----------------------------|
| Monday    | Rosanna      | 300      | 1 header + up to 300 rows |
| Tuesday   | Joshua       | 300      | 1 header + up to 300 rows |
| Wednesday | —            | 0 (none) | Not generated              |
| Thursday  | Joshua       | 125      | 1 header + up to 125 rows |
| Friday    | Joshua       | 125      | 1 header + up to 125 rows |
| Sat/Sun   | —            | 0 (none) | Not generated              |

Everything past the capped staff member's share of the professional pool goes to
Jasmine, along with any billable Programming/e-care rows for that day.

The system recognizes e-care variants 'e-care', 'e care', 'ecare', and
'extended care' (case-insensitive).

**Melissa** and the O'Flynn Karen "Unable to Bill" rule take priority over the
Rosanna/Joshua/Jasmine schedule above:
- WM/OP WM Program Level rows always go to Melissa.
- PHP/Partial Hospitalization rows always go to Melissa, every day. She does not
  get an individual report — PHP rows are assigned to her in the Masters
  spreadsheet only. PHP is billed only on Tuesdays as an operational matter.
- Detox/Residential rows billed to Aetna or Humana (and not a drug screen) go to
  Melissa.
- Billing Provider "O'Flynn, Karen" with GROUPFLD1 "OP Chappaqua" or "OP NYC" is
  always Unable to Bill.

## Reports

Individual workbooks are generated for **Rosanna**, **Joshua**, **Jasmine**, and
**CB** (empty reports are skipped, e.g. Rosanna on any day but Monday, or Joshua
on Monday/Wednesday/weekends). All other staff (Melissa, Unable to Bill, etc.)
are still assigned in the Masters workbook but do not receive separate reports.

Rosanna's, Joshua's, and Jasmine's reports include a Status column with a
dropdown list: Billed, Unable to Bill, Contractual Adj, Incomplete Billings,
Utox Batch, and Inclusive Services. Jasmine's report also includes two
Jasmine-only options: Batch Billings and IOP.

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
