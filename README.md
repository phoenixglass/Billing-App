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

**Insurance** (`GROUPFLD2` = "Insurance") rows are split between **Rosanna** and
**Jasmine**. `GROUPFLD2` values other than "Insurance" or "Self Pay" never
reach Rosanna or Jasmine — they are marked Unable to Bill.

- **Professional** services (identified by the `Claim Type` column equal to
  `CMS-1500` or `UB-04` — UB-04 counts as Professional every day) bill every
  day of the week. The Insurance + CMS-1500/UB-04 rows ("the professional
  pool") are sorted alphabetically by `Client`. Monday through Friday,
  Rosanna receives the first 150 rows of that sorted pool; the rest of the
  pool goes to Jasmine. Rosanna caps no rows on weekends, so Jasmine gets
  the whole pool those days.
- **Programming** services (Detox, Residential) bill Tuesday, Thursday,
  Friday, and weekends; they are Unable to Bill on Monday and Wednesday. All
  billable Programming rows go to Jasmine.
- **IOP** (including Telemed IOP) bills every day of the week, with no
  exceptions, and always goes to Jasmine, bypassing the professional
  pool/Rosanna split even if Claim Type is CMS-1500 or UB-04 — unless the
  Cathy report is on and the row is a Professional row for one of her
  payers, which is hers (see below).
- **E-care** bills on Tuesdays only (regardless of Claim Type). Billable e-care
  rows go to Jasmine.
- Any other billable Insurance row whose Claim Type is not CMS-1500/UB-04
  (i.e. institutional/837I) goes to Jasmine, unless it's PHP (always Melissa's).
- **PHP** (Partial Hospitalization) always goes to **Melissa**, every day — see
  the Melissa section below. It is not part of the Programming bucket above and
  never reaches Rosanna or Jasmine.

Rosanna's professional-pool cap by weekday:

| Day       | Capped staff | Cap      | Report                    |
|-----------|--------------|----------|----------------------------|
| Monday    | Rosanna      | 150      | 1 header + up to 150 rows |
| Tuesday   | Rosanna      | 150      | 1 header + up to 150 rows |
| Wednesday | Rosanna      | 150      | 1 header + up to 150 rows |
| Thursday  | Rosanna      | 150      | 1 header + up to 150 rows |
| Friday    | Rosanna      | 150      | 1 header + up to 150 rows |
| Sat/Sun   | —            | 0 (none) | Not generated              |

Everything past Rosanna's share of the professional pool goes to Jasmine,
along with any billable Programming/e-care rows for that day. The per-run
"Don't give Rosanna anything" option drops her cap to zero for every day, so
Jasmine takes the whole pool.

The system recognizes e-care variants 'e-care', 'e care', 'ecare', and
'extended care' (case-insensitive).

**Melissa** and the O'Flynn Karen "Unable to Bill" rule take priority over the
Rosanna/Jasmine schedule above:
- WM/OP WM Program Level rows always go to Melissa.
- PHP/Partial Hospitalization rows always go to Melissa, every day. She does not
  get an individual report — PHP rows are assigned to her in the Masters
  spreadsheet only. PHP is billed only on Tuesdays as an operational matter.
- Detox/Residential rows billed to Aetna or Humana (and not a drug screen) go to
  Melissa.
- Billing Provider "O'Flynn, Karen" with GROUPFLD1 "OP Chappaqua" or "OP NYC" is
  always Unable to Bill.

## Optional Per-Run Options

The checkboxes in the app are per-run options for the file being processed.
All of them are off by default, so an unchecked run follows the standard daily
schedule above. The command-line script takes the same options as flags.

| Option | Flag | Effect |
|--------|------|--------|
| Exclude Optum insurance | — | Optum utox (drug screen) rows are left out of the individual workbooks. |
| Exclude BCB Anthem CT for PHP, Residential, and Detox | — | Those BCB Anthem CT rows are left out of the individual workbooks. |
| Remove Anthem from Rosanna and Jasmine reports | — | Anthem rows are left out of Rosanna's and Jasmine's workbooks. |
| Don't give anyone Detox or Residential | — | Detox/Residential rows are left out of every individual workbook. |
| Include Programming (Detox/Residential) today | `--include-programming` | Programming bills regardless of the weekday, so it can be worked on a Monday or Wednesday. Billable Programming rows go to Jasmine as usual. E-care is unaffected and stays Tuesday-only. |
| Exclude Aetna | `--exclude-aetna` | Every Aetna row is left out of the individual workbooks. |
| Cathy report: Professional services only for Oxford, ConnectiCare, UBH | `--cathy-report` | See the Cathy section below. |
| Cathy report: all of her payers (ConnectiCare, Emblem, Oxford, Surest, UBH, UBH-HP, UMR) | `--cathy-all-payers` | The same Cathy report run against her full payer list instead of just her usual three. Turns the Cathy report on by itself — the box above does not also need to be checked. See the Cathy section below. |
| Don't give Rosanna anything | `--no-rosanna` | Rosanna is assigned no rows and gets no workbook; her share of the professional pool goes to Jasmine, the same way it does on a weekend. |

Rows excluded by any of these options are still assigned in the Masters
workbook — the option only controls what reaches the individual reports.

Checking both "Include Programming (Detox/Residential) today" and "Don't give
anyone Detox or Residential" is contradictory; the exclusion wins, and the app
shows a warning saying so.

### Cathy (optional report)

When the Cathy report is turned on, **every** Insurance row whose `Claim Type`
is Professional (`CMS-1500` or `UB-04`) **and** whose `Payer` is on her payer
list is assigned to **Cathy** and saved as her own workbook. There are two
payer lists to choose from:

| Option | Payers |
|--------|--------|
| Cathy report: Professional services only for Oxford, ConnectiCare, UBH | Oxford, ConnectiCare, UBH (UBH-HP included — it matches the UBH pattern) |
| Cathy report: all of her payers | The three above plus Emblem, Surest, UMR — i.e. ConnectiCare, Emblem, Oxford, Surest, UBH, UBH-HP, UMR |

The wider list changes **only** which payers are hers; everything else about
the report is the same, and checking it runs the Cathy report on its own
whether or not the narrower box is also checked.

- The service does not matter, only the claim type and the payer. IOP for
  her payers is hers too: she takes it ahead of the IOP-to-Jasmine
  rule. IOP for any other payer, or IOP that is not a Professional claim
  type, is still Jasmine's.
- Those rows leave the Rosanna/Jasmine professional pool rather than being
  duplicated into it, so no row is worked twice. Rosanna's 150-row cap then
  applies to whatever is left of the pool.
- Payer matching is case-insensitive and tolerates the spelling variants these
  payers appear with: `ConnectiCare`/`Connecti Care`, `UBH`/`United
  Behavioral Health`, and the `(Optum)` suffixes (`Emblem (Optum)`,
  `Surest (Optum)`, `UBH-HP (Optum)`, `UMR (Optum)`). `UBH` and `UMR` only
  match as whole words, so they are not picked up inside a longer word.
- Three rules still take priority over Cathy, even for her own payers:
  WM/OP WM (Melissa — only she is authorized to bill WM), PHP (Melissa), and
  Billing Provider "O'Flynn, Karen" in OP Chappaqua/OP NYC (always Unable to
  Bill). Self Pay rows still go to CB.
- Her workbook gets Status/Comments columns whose dropdown carries the same
  options as Jasmine's, Batch Billings and IOP included.

### Giving Rosanna nothing (optional)

When "Don't give Rosanna anything" is turned on, Rosanna is assigned no rows
for that run and no workbook is generated for her. Her share of the
professional pool goes to **Jasmine** instead — the same thing that already
happens on a weekend, when Rosanna's cap is zero. Nothing is left unassigned:
every row still appears in the Masters workbook with an owner, and the rules
that never involved Rosanna (Self Pay to CB, Melissa's rows, Cathy's rows when
her report is on) are untouched.

## Reports

Individual workbooks are generated for **Rosanna**, **Jasmine**, and **CB**
(empty reports are skipped, e.g. Rosanna on weekends, and Rosanna's is not
generated at all when "Don't give Rosanna anything" is on), plus **Cathy** when
her report is turned on for the run. All other staff (Melissa, Unable to Bill,
etc.) are still assigned in the Masters workbook but do not receive separate
reports.

Rosanna's, Jasmine's, and Cathy's reports include a Status column with a
dropdown list: Billed, Unable to Bill, Contractual Adj, Incomplete Billings,
Utox Batch, and Inclusive Services. Jasmine's and Cathy's reports share the
same dropdown, which adds two more options: Batch Billings and IOP.

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
python "Unbilled Step 1.py" "path/to/Unbilled Revenue 01252026.xlsx"
```

Add any of the per-run flags as needed, for example:

```bash
python "Unbilled Step 1.py" "path/to/file.xlsx" --include-programming --exclude-aetna --cathy-report

# Cathy's full payer list, and nothing for Rosanna:
python "Unbilled Step 1.py" "path/to/file.xlsx" --cathy-all-payers --no-rosanna
```

## Testing

Run the unit tests for weekday rules:

```bash
python tests/test_weekday_rules.py
python tests/test_assign_staff.py
```

Or use pytest if available:

```bash
pytest tests/ -v
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
- `tests/test_assign_staff.py` - End-to-end tests for staff assignment
- `requirements.txt` - Python dependencies

## License

Proprietary
