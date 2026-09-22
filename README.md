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

**Insurance** (`GROUPFLD2` = "Insurance") rows are split between **Cathy** and
**Jasmine**. `GROUPFLD2` values other than "Insurance" or "Self Pay" never
reach Cathy or Jasmine — they are marked Unable to Bill. Cathy fills the role
Rosanna used to hold, on top of her own payer-specific carve-out (see the
Cathy section below) — the two stack and never double-assign the same row.

- **Professional** services (identified by the `Claim Type` column equal to
  `CMS-1500` or `UB-04` — UB-04 counts as Professional every day) bill every
  day of the week. The Insurance + CMS-1500/UB-04 rows ("the professional
  pool") are sorted alphabetically by `Client`. Monday through Friday,
  Cathy receives the first 150 rows of that sorted pool; the rest of the
  pool goes to Jasmine. Cathy caps no rows on weekends, so Jasmine gets
  the whole pool those days.
- **Programming** services (Detox, Residential) bill Tuesday, Thursday,
  Friday, and weekends; they are Unable to Bill on Monday and Wednesday. All
  billable Programming rows go to Jasmine.
- **IOP** (including Telemed IOP) bills every day of the week, with no
  exceptions, and always goes to Jasmine, bypassing the professional
  pool/Cathy split even if Claim Type is CMS-1500 or UB-04 — unless her
  payer carve-out is on and the row is a Professional row for one of her
  payers, which is hers (see below).
- **E-care** bills on Tuesdays only (regardless of Claim Type). Billable e-care
  rows go to Jasmine.
- Any other billable Insurance row whose Claim Type is not CMS-1500/UB-04
  (i.e. institutional/837I) goes to Jasmine, unless it's PHP (always Melissa's).
- **PHP** (Partial Hospitalization) always goes to **Melissa**, every day — see
  the Melissa section below. It is not part of the Programming bucket above and
  never reaches Cathy or Jasmine.

Cathy's professional-pool cap by weekday:

| Day       | Capped staff | Cap      | Report                    |
|-----------|--------------|----------|----------------------------|
| Monday    | Cathy        | 150      | 1 header + up to 150 rows |
| Tuesday   | Cathy        | 150      | 1 header + up to 150 rows |
| Wednesday | Cathy        | 150      | 1 header + up to 150 rows |
| Thursday  | Cathy        | 150      | 1 header + up to 150 rows |
| Friday    | Cathy        | 150      | 1 header + up to 150 rows |
| Sat/Sun   | —            | 0 (none) | Not generated              |

Everything past Cathy's share of the professional pool goes to Jasmine,
along with any billable Programming/e-care rows for that day. The per-run
"Don't give Cathy anything" option drops her cap to zero for every day (and
turns off her payer carve-out too), so Jasmine takes the whole pool. The
per-run "Split all services evenly between Jasmine and Cathy" option divides
everything Jasmine would otherwise get 50/50 with Cathy instead — see
"Splitting services evenly" below.

The system recognizes e-care variants 'e-care', 'e care', 'ecare', and
'extended care' (case-insensitive).

**Melissa** and the O'Flynn Karen "Unable to Bill" rule take priority over the
Cathy/Jasmine schedule above:
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
| Remove Anthem from Cathy and Jasmine reports | — | Anthem rows are left out of Cathy's and Jasmine's workbooks. |
| Don't give anyone Detox or Residential | — | Detox/Residential rows are left out of every individual workbook. |
| Include Programming (Detox/Residential) today | `--include-programming` | Programming bills regardless of the weekday, so it can be worked on a Monday or Wednesday. Billable Programming rows go to Jasmine as usual. E-care is unaffected and stays Tuesday-only. |
| Exclude Aetna | `--exclude-aetna` | Every Aetna row is left out of the individual workbooks. |
| Cathy carve-out: Professional services only for Oxford, ConnectiCare, UBH | `--cathy-report` | See the Cathy section below. |
| Cathy carve-out: all of her payers (ConnectiCare, Emblem, Oxford, Surest, UBH, UBH-HP, UMR) | `--cathy-all-payers` | The same carve-out run against her full payer list instead of just her usual three. Turns the carve-out on by itself — the box above does not also need to be checked. See the Cathy section below. |
| Don't give Cathy anything | `--no-cathy` | Cathy is assigned no rows at all and gets no workbook — neither her carve-out nor her pool share; her share of the professional pool goes to Jasmine, the same way it does on a weekend. |
| Split all services evenly between Jasmine and Cathy | `--even-split-jasmine-cathy` | Everything Jasmine would otherwise get (the rest of the pool, plus billable Programming/e-care, plus other billable institutional rows) is split 50/50 with Cathy instead. See "Splitting services evenly" below. |

Rows excluded by any of these options are still assigned in the Masters
workbook — the option only controls what reaches the individual reports.

### Custom exclusions (free text, no code change needed)

Alongside the fixed checkboxes above, the app has two free-text fields for
one-off exclusions that don't have a checkbox yet:

- **Exclude payers containing** — a comma-separated list of terms
  (case-insensitive substring match against the Payer column). Example:
  `Cigna, Humana`.
- **Exclude services containing** — the same, matched against the Service
  column. Example: `Group Therapy`.
- **Apply only to these staff** — an optional list limiting the two fields
  above to specific staff workbooks (Jasmine, Cathy, CB). Leave it
  empty to apply them to every individual workbook, the same way Exclude
  Aetna does.

These behave exactly like the checkboxes: matching rows are left out of the
individual workbooks for that run only, and still appear in the Masters
workbook. They exist so a payer or service that comes up once — "leave out
Cigna today" — doesn't need a new checkbox, a code change, and a redeploy;
type it into the box and process the file. A rule that turns out to be
needed every time is still a good candidate to become a real checkbox later.

The command-line script takes the same fields as flags: `--exclude-payers`,
`--exclude-services`, and `--exclude-scope` (all comma-separated).

Checking both "Include Programming (Detox/Residential) today" and "Don't give
anyone Detox or Residential" is contradictory; the exclusion wins, and the app
shows a warning saying so.

### Removing a funding source by division (optional, no code change needed)

"Remove a funding source by division" excludes a payer, but only within
specific divisions, instead of everywhere. Two comma-separated fields, both
supporting **multiple values**:

- **Funding source(s) to remove** — case-insensitive substring match against
  the Payer column. Example: `BCBS, Beacon`.
- **Division(s) to remove it from** — case-insensitive substring match
  against the `GROUPFLD1` column (e.g. Residential, Detox, OP Wilton, OP
  Canaan). Example: `Residential, Detox, OP Wilton, OP Canaan`.

A row is excluded from every individual workbook only when its Payer matches
**one of** the funding source(s) **and** its GROUPFLD1 matches **one of** the
division(s) — every combination of the two lists is removed, so entering
several funding sources and several divisions in the same run covers all of
them together, not just one pair at a time. Leaving either field blank (or
entering only one of the two) does nothing; the app warns if that happens.
Rows still appear in the Masters workbook.

The command-line script takes the same fields as
`--exclude-payer-by-division-payers` and
`--exclude-payer-by-division-divisions` (both comma-separated).

### Cathy (payer carve-out and standing pool share)

Cathy has two ways to receive rows, and they stack rather than replace one
another:

1. **Her standing share of the professional pool** — the role Rosanna used
   to hold (see "Daily Billing Rules" above): up to 150 Professional rows a
   weekday, sorted alphabetically, with the rest going to Jasmine.
2. **Her payer carve-out (optional)** — when turned on, **every** Insurance
   row whose `Claim Type` is Professional (`CMS-1500` or `UB-04`) **and**
   whose `Payer` is on her payer list is assigned to her regardless of the
   pool/cap, and leaves the pool entirely so it's never worked twice. There
   are two payer lists to choose from:

| Option | Payers |
|--------|--------|
| Cathy carve-out: Professional services only for Oxford, ConnectiCare, UBH | Oxford, ConnectiCare, UBH (UBH-HP included — it matches the UBH pattern) |
| Cathy carve-out: all of her payers | The three above plus Emblem, Surest, UMR — i.e. ConnectiCare, Emblem, Oxford, Surest, UBH, UBH-HP, UMR |

The wider list changes **only** which payers are hers; everything else about
the carve-out is the same, and checking it runs the carve-out on its own
whether or not the narrower box is also checked.

- The service does not matter, only the claim type and the payer. IOP for
  her carve-out payers is hers too: she takes it ahead of the IOP-to-Jasmine
  rule. IOP for any other payer, or IOP that is not a Professional claim
  type, is still Jasmine's.
- Carve-out rows leave the Cathy/Jasmine professional pool rather than being
  duplicated into it, so no row is worked twice. Her 150-row cap (or the
  even split, if that's on) then applies to whatever is left of the pool.
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

### Giving Cathy nothing (optional)

When "Don't give Cathy anything" is turned on, Cathy is assigned no rows at
all for that run — neither her payer carve-out nor her share of the
professional pool — and no workbook is generated for her. Her share of the
pool goes to **Jasmine** instead — the same thing that already happens on a
weekend, when her cap is zero. Nothing is left unassigned: every row still
appears in the Masters workbook with an owner, and the rules that never
involved Cathy (Self Pay to CB, Melissa's rows) are untouched.

### Overriding Cathy's cap (optional, no code change needed)

"Override Cathy's cap for today" replaces the standard weekday schedule (150
Monday-Friday, 0 on weekends) with an exact row count for this run only — for
example, giving her 100 on a weekday she's out for part of, or opening up 20
rows for her on a weekend. It's ignored if "Don't give Cathy anything" or
"Split all services evenly between Jasmine and Cathy" is also checked. The
command-line script takes the same option as `--cathy-cap N`.

### Splitting services evenly (optional)

"Split all services evenly between Jasmine and Cathy" replaces Cathy's cap
entirely for the run: instead of her taking a fixed row count and Jasmine
the remainder, **everything Jasmine would otherwise receive** is divided
50/50 between the two — the rest of the professional pool, plus billable
Programming/e-care, plus other billable non-Professional (institutional)
Insurance rows. IOP still always goes to Jasmine no matter what. Cathy's
payer carve-out, if also turned on, still claims its rows first, ahead of
the split. It's ignored if "Don't give Cathy anything" is also checked (she
still gets nothing). The command-line script takes the same option as
`--even-split-jasmine-cathy`.

### Custom report (optional, no code change needed)

The "Custom report" fields are a second, generic version of Cathy's carve-out
above, for routing a specific payer's rows to a **different** staff member
without a checkbox or a code change:

- **Staff name for this report** — who the matching rows go to. Must not be
  one of the reserved names (Jasmine, CB, Melissa, Cathy, Unable to
  Bill); the app rejects the run with an error if it collides.
- **Payers for this report** — comma-separated, case-insensitive substring
  match against the Payer column, same matching as the custom exclusion
  fields.
- **Professional claim types only** — checked by default (the same
  restriction Cathy has: only CMS-1500/UB-04 rows for these payers are
  claimed). Uncheck to match any claim type.

Like Cathy's carve-out, this report's rows leave the Cathy/Jasmine
professional pool entirely (no row is worked twice), and it's checked *after*
Cathy's carve-out, so if a payer is on both lists, Cathy's rows stay hers.
The rules that outrank Cathy (WM/OP WM, PHP, the O'Flynn Karen rule, Self Pay
to CB) outrank this report too. Its workbook gets the same Status dropdown as
Jasmine's and Cathy's (Batch Billings and IOP included). Both fields must be
set for the report to run — a name with no payers, or payers with no name,
do nothing.

The command-line script takes the same options as `--custom-report-name`,
`--custom-report-payers`, and `--custom-report-any-claim-type` (to turn off
the Professional-only restriction).

## Reports

Individual workbooks are generated for **Cathy**, **Jasmine**, and **CB**
(empty reports are skipped, e.g. Cathy on weekends, and Cathy's is not
generated at all when "Don't give Cathy anything" is on), plus the
**custom report**'s staff member when that's configured. All other staff
(Melissa, Unable to Bill, etc.) are still assigned in the Masters workbook
but do not receive separate reports.

Jasmine's, Cathy's, and the custom report's workbooks share a Status column
with a dropdown list: Billed, Unable to Bill, Contractual Adj, Incomplete
Billings, Utox Batch, Inclusive Services, Batch Billings, and IOP.

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

# Cathy's full payer list, and nothing for Cathy otherwise:
python "Unbilled Step 1.py" "path/to/file.xlsx" --cathy-all-payers --no-cathy

# Split everything else 50/50 between Jasmine and Cathy, and remove BCBS/Beacon
# from the Residential and Detox divisions:
python "Unbilled Step 1.py" "path/to/file.xlsx" --even-split-jasmine-cathy \
    --exclude-payer-by-division-payers "BCBS, Beacon" \
    --exclude-payer-by-division-divisions "Residential, Detox"
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

- `app.py` - Streamlit web application (UI only; the assignment logic itself
  lives in `billing_rules.py`)
- `Unbilled Step 1.py` - Command-line processing script (same relationship
  to `billing_rules.py` as app.py)
- `billing_rules.py` - The shared assignment engine: `assign_staff`,
  `finalize_workbook`, the weekday/payer/claim-type classifiers, and the
  free-text exclusion/custom-report helpers. app.py and
  `Unbilled Step 1.py` both import from here rather than keeping their own
  copies, so a rule change can't happen in one and not the other.
- `tests/test_weekday_rules.py` - Unit tests for the classifier helpers
- `tests/test_assign_staff.py` - End-to-end tests for staff assignment,
  loaded against `Unbilled Step 1.py` (which just re-exports
  `billing_rules.assign_staff`)
- `requirements.txt` - Python dependencies

## License

Proprietary
