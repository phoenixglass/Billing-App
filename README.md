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
reach Cathy or Jasmine — they are marked Unable to Bill.

**Every day of the week, everything that goes to Cathy or Jasmine is split
exactly in half between them.** That "shared pool" is every Insurance row
either of them would get:

- **Professional** services (Claim Type `CMS-1500` or `UB-04` — UB-04
  counts as Professional every day), which bill every day of the week.
- **IOP** (including Telemed IOP), which bills every day of the week.
- **Programming** (Detox, Residential), which bills Tuesday, Thursday,
  Friday, and weekends; it is Unable to Bill on Monday and Wednesday.
- **E-care**, which bills on Tuesdays only (regardless of Claim Type).
- Any other billable Insurance row whose Claim Type is not CMS-1500/UB-04
  (i.e. institutional/837I).

**PHP** (Partial Hospitalization) is not in the pool: all Insurance PHP goes
to **Melissa**, every day (see the Melissa section below).

### How the split works: Day A and Day B

The shared pool is sorted alphabetically by `Client` and cut at its exact
midpoint into a first half (the A-M side) and a second half (the N-Z side).
Which half each person gets alternates every day:

| Day   | Jasmine gets              | Cathy gets                |
|-------|---------------------------|---------------------------|
| Day A | First half (clients A-M)  | Second half (clients N-Z) |
| Day B | Second half (clients N-Z) | First half (clients A-M)  |

The days run A, B, A, B, … so whoever had A-M one day has N-Z the next.

- **Which day is which** is worked out automatically from the date in the
  filename: 09/23/2026 is Day A, and each calendar day after that flips
  (09/24/2026 is Day B, 09/25/2026 is Day A, and so on). Weekends count, so
  the rotation stays in step even when a day is skipped. The Day A anchor
  date is `SPLIT_DAY_A_ANCHOR` in `billing_rules.py`.
- **Forcing a day:** the app's "Jasmine/Cathy split day" dropdown (or the
  CLI's `--split-day A|B`) overrides the automatic day for one run, in
  case the rotation ever needs correcting. After processing, the app shows
  which day it used.
- **"Exactly half" is counted in rows, not letters.** The cut sits at the
  midpoint of the sorted pool, so the two halves are equal even when most
  clients' names fall in A-M. "A-M" and "N-Z" describe which end of the
  alphabet each half comes from; the actual boundary is wherever the
  midpoint lands (for example, Jasmine might get A through "Lopez" and
  Cathy "Lopez" through Z). A client with several rows can therefore be
  split across the two halves if their rows sit right at the midpoint.
- **Odd row counts:** an exact half isn't possible, so the second (N-Z)
  half gets the one extra row.
- **Exclusions are applied afterwards.** The per-run exclusions below
  (Exclude Aetna, Remove Anthem, etc.) remove rows from the individual
  workbooks after the split, so if they remove more rows from one half
  than the other, the two workbooks can end up uneven.

The system recognizes e-care variants 'e-care', 'e care', 'ecare', and
'extended care' (case-insensitive).

**Melissa** and the O'Flynn Karen "Unable to Bill" rule take priority over the
Cathy/Jasmine schedule above:
- WM/OP WM Program Level rows always go to Melissa.
- All Insurance PHP/Partial Hospitalization rows go to Melissa, every day,
  whatever the payer or claim type — ahead of the O'Flynn Karen rule below,
  too. She does not get an individual report — PHP rows are assigned to her
  in the Masters spreadsheet only. PHP is billed only on Tuesdays as an
  operational matter. (Self Pay PHP still goes to CB, like all Self Pay.)
- Detox/Residential rows billed to Aetna or Humana (and not a drug screen) go to
  Melissa.
- Billing Provider "O'Flynn, Karen" with GROUPFLD1 "OP Chappaqua" or "OP NYC" is
  Unable to Bill (other than PHP, which is Melissa's).

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
| Include Programming (Detox/Residential) today | `--include-programming` | Programming bills regardless of the weekday, so it can be worked on a Monday or Wednesday. Billable Programming rows are split between Jasmine and Cathy as usual. E-care is unaffected and stays Tuesday-only. |
| Exclude Aetna | `--exclude-aetna` | Every Aetna row is left out of the individual workbooks. |
| Cathy carve-out: Professional services only for Oxford, ConnectiCare, UBH | `--cathy-report` | See the Cathy section below. |
| Cathy carve-out: all of her payers (ConnectiCare, Emblem, Oxford, Surest, UBH, UBH-HP, UMR) | `--cathy-all-payers` | The same carve-out run against her full payer list instead of just her usual three. Turns the carve-out on by itself — the box above does not also need to be checked. See the Cathy section below. |
| Don't give Cathy anything | `--no-cathy` | Cathy is assigned no rows at all and gets no workbook — neither her carve-out nor her half of the shared pool; Jasmine gets the whole pool. |
| Jasmine/Cathy split day (dropdown: Automatic / Day A / Day B) | `--split-day A` or `--split-day B` | Forces that day's split for this run instead of working it out from the filename's date. See "How the split works" above. |

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

1. **Her half of the shared pool** (see "Daily Billing Rules" above): the
   A-M half on Day B, the N-Z half on Day A.
2. **Her payer carve-out (optional)** — when turned on, **every** Insurance
   row whose `Claim Type` is Professional (`CMS-1500` or `UB-04`) **and**
   whose `Payer` is on her payer list is assigned to her outright, and
   leaves the shared pool entirely so it's never worked twice. There are
   two payer lists to choose from:

| Option | Payers |
|--------|--------|
| Cathy carve-out: Professional services only for Oxford, ConnectiCare, UBH | Oxford, ConnectiCare, UBH (UBH-HP included — it matches the UBH pattern) |
| Cathy carve-out: all of her payers | The three above plus Emblem, Surest, UMR — i.e. ConnectiCare, Emblem, Oxford, Surest, UBH, UBH-HP, UMR |

The wider list changes **only** which payers are hers; everything else about
the carve-out is the same, and checking it runs the carve-out on its own
whether or not the narrower box is also checked.

- The service does not matter, only the claim type and the payer. IOP for
  her carve-out payers is hers too. IOP for any other payer, or IOP that is
  not a Professional claim type, stays in the shared pool.
- Carve-out rows leave the shared pool rather than being duplicated into
  it, so no row is worked twice. The daily half-and-half split then applies
  to whatever is left of the pool — so **with the carve-out on, Cathy ends
  up with more than half overall** (her carve-out rows plus half of the
  rest).
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
all for that run — neither her payer carve-out nor her half of the shared
pool — and no workbook is generated for her. **Jasmine** gets the whole
pool instead. Nothing is left unassigned: every row still appears in the
Masters workbook with an owner, and the rules that never involved Cathy
(Self Pay to CB, Melissa's rows) are untouched. The Day A/Day B rotation
keeps counting on days Cathy is skipped, so the next day's split is the
same as if she hadn't been.

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

Like Cathy's carve-out, this report's rows leave the shared Cathy/Jasmine
pool entirely (no row is worked twice), and it's checked *after*
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
(empty reports are skipped, and Cathy's is not generated at all when
"Don't give Cathy anything" is on), plus the
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

# Force today's Jasmine/Cathy split to Day B, and remove BCBS/Beacon from the
# Residential and Detox divisions:
python "Unbilled Step 1.py" "path/to/file.xlsx" --split-day B \
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
