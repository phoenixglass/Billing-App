"""
Helper functions to:
- parse MMDDYYYY date tokens into weekday
- determine whether a given service is non-billable for a given weekday
- classify a "Claim Type" value as Professional (CMS-1500 or UB-04)

Weekday mapping: 0=Monday, 1=Tuesday, ..., 6=Sunday

Daily billing schedule:
- Professional services (Claim Type "CMS-1500" or "UB-04" — UB-04 counts as
  Professional every day of the week) bill every day of the week.
- IOP (including Telemed IOP) bills every day of the week, with no
  exceptions.
- Programming (Detox, Residential) bills Tuesday, Thursday, Friday, and
  weekends; non-billable Monday and Wednesday.
- E-care bills on Tuesdays only, regardless of Claim Type.
- Self Pay: every service bills every day, with no exceptions (including
  e-care).
- PHP (Partial Hospitalization) is not part of the Programming schedule
  above: it always goes to Melissa (see is_php_service), and is billed
  only on Tuesdays as a real-world/manual matter, not something this
  module's weekday helpers gate.

Cathy/Jasmine split (applies to GROUPFLD2 == "Insurance" rows only; no
Self Pay ever goes to either of them, and PHP never reaches them since it
goes to Melissa). Every day of the week, everything that
goes to Cathy or Jasmine is split exactly in half between the two:
- The "shared pool" is every Insurance row either of them would get:
  Professional rows (Claim Type CMS-1500 or UB-04), IOP (including
  Telemed IOP), billable Programming/e-care, and any other billable
  Insurance row whose Claim Type is not CMS-1500/UB-04 (i.e.
  institutional/837I). It is sorted alphabetically by Client and cut
  at its exact midpoint into a first (A-M side) and a second (N-Z side)
  half. With an odd row count the second half gets the one extra row.
- The halves alternate daily between "Day A" and "Day B" (see
  split_day_for_date_token): on Day A Jasmine gets the first (A-M) half
  and Cathy the second (N-Z) half; on Day B they swap. Consecutive
  calendar days always alternate, weekends included, and
  SPLIT_DAY_A_ANCHOR is a known Day A.
- GROUPFLD2 values other than "Insurance" or "Self Pay" never reach
  Cathy or Jasmine.

Optional, per-run overrides (all off by default; nothing below changes the
standard schedule unless the operator turns it on for that run):
- include_programming: bill Programming (Detox, Residential) regardless of
  the weekday, so it can be included on a Monday or Wednesday. E-care is
  unaffected and stays Tuesday-only.
- Nothing for Cathy: give Cathy no rows at all for the run. The whole
  shared pool goes to Jasmine instead, and no workbook is generated for
  her.
- split_day: force "A" or "B" for this run instead of deriving it from
  the file's date, in case the alternation ever needs correcting.
- Report exclusions (Exclude Aetna, Remove Anthem, the free-text payer/
  service/division exclusions, etc. — see ReportExclusions): drop rows
  from the individual staff reports; they stay in the Masters workbook.
  They are applied before the Jasmine/Cathy split, so the rows that
  actually reach the two reports are what gets divided in half.
"""
from datetime import date, datetime
from typing import Tuple

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


DRUG_SCREEN_KEYWORDS = ("drug screen", "utox", "urine tox", "drug test", "uds")
PHP_KEYWORDS = ("partial hospitalization", "php")

# A known "Day A" for the daily Jasmine/Cathy split: on Day A Jasmine gets
# the first (A-M) half of the alphabetically sorted shared pool and Cathy
# the second (N-Z) half; on Day B they swap. Every calendar day alternates,
# so any date an even number of days from this one is Day A and any date
# an odd number of days away is Day B. 09/23/2026 is the day the
# alternating split started; if the rotation ever needs to be shifted by a
# day, change this date (or force the day for one run via split_day).
SPLIT_DAY_A_ANCHOR = date(2026, 9, 23)
SPLIT_DAYS = ("A", "B")

def parse_weekday_from_token(date_token: str) -> Tuple[int, bool]:
    """
    Parse an MMDDYYYY date_token and return (weekday, did_fallback).
    - weekday: integer 0..6
    - did_fallback: True if parsing failed and we fell back to today's weekday
    """
    if not date_token:
        return datetime.now().weekday(), True
    try:
        dt = datetime.strptime(date_token, "%m%d%Y")
        return dt.weekday(), False
    except Exception:
        return datetime.now().weekday(), True


def split_day_for_date_token(date_token: str) -> Tuple[str, bool]:
    """Return ("A" or "B", did_fallback) for the Jasmine/Cathy split.

    Day A: Jasmine gets the first (A-M) half of the shared pool, Cathy the
    second (N-Z) half. Day B: the reverse. The day is derived from the
    MMDDYYYY date_token by counting calendar days from SPLIT_DAY_A_ANCHOR,
    so it alternates every day, weekends included, and it doesn't depend
    on which days a file actually gets run. Falls back to today's date if
    the token is missing or unparseable, the same way
    parse_weekday_from_token does.
    """
    did_fallback = False
    try:
        day = datetime.strptime(date_token, "%m%d%Y").date()
    except (TypeError, ValueError):
        day = datetime.now().date()
        did_fallback = True
    return SPLIT_DAYS[(day - SPLIT_DAY_A_ANCHOR).days % 2], did_fallback


def _is_ecare(service: str) -> bool:
    """Return True if the service text refers to e-care (many possible variants)."""
    s = service.lower()
    return any(token in s for token in ("e-care", "e care", "ecare", "extended care"))


def _is_drug_screen(service: str) -> bool:
    """Return True if the service text refers to a drug screen (Utox)."""
    s = service.lower()
    return any(kw in s for kw in DRUG_SCREEN_KEYWORDS)


def is_php_service(service: str) -> bool:
    """Return True for PHP/Partial Hospitalization services (always Melissa's)."""
    s = (service or "").lower()
    return any(kw in s for kw in PHP_KEYWORDS)


def _is_programming_service(service: str) -> bool:
    """Return True for Programming services: Detox, Residential.

    IOP is intentionally excluded here: it bills every day of the week
    (see is_iop_service) rather than following this weekday schedule.
    PHP is also excluded: it's always assigned to Melissa (see
    is_php_service) rather than following this weekday schedule.
    """
    s = service.lower()
    return (
        "detox" in s
        or "residential" in s
    )


def is_iop_service(service: str) -> bool:
    """Return True for any IOP service, including Telemed IOP (case-insensitive)."""
    return "iop" in (service or "").lower()


def is_professional_claim_type(claim_type: str) -> bool:
    """Return True if the Claim Type column value is CMS-1500 or UB-04 (Professional).

    UB-04 counts as Professional every day of the week, alongside CMS-1500.
    """
    return (claim_type or "").strip().upper() in ("CMS-1500", "UB-04")


def is_aetna_payer(payer: str) -> bool:
    """Return True if the Payer column value refers to Aetna."""
    return "aetna" in (payer or "").lower()


def parse_terms(raw: str) -> list:
    """Split a comma-separated string into trimmed, non-empty terms.

    Used by the custom per-run payer/service exclusion fields, so an
    operator can type "Cigna, Humana" into a text box instead of a new
    checkbox needing a code change for each new payer or service.
    """
    if not raw:
        return []
    return [term.strip() for term in raw.split(",") if term.strip()]


def matches_any_term(text: str, terms) -> bool:
    """Return True if text contains any of terms, case-insensitively."""
    if not terms:
        return False
    haystack = (text or "").lower()
    return any(term.lower() in haystack for term in terms)


def payer_excluded_by_division(payer: str, division: str, payer_terms, division_terms) -> bool:
    """Return True if payer matches payer_terms AND division matches division_terms.

    Used by the optional "remove a funding source by division" exclusion:
    an operator names a payer (e.g. BCBS, Beacon) and a division (the
    GROUPFLD1 column, e.g. Residential, Detox, OP Wilton, OP Canaan), and
    only rows matching both are excluded. Both term lists must be
    non-empty, or this never matches — entering only one field does
    nothing, since a payer alone or a division alone isn't the feature
    being asked for.
    """
    if not payer_terms or not division_terms:
        return False
    return matches_any_term(payer, payer_terms) and matches_any_term(division, division_terms)


def is_wm_program_level(cell_value) -> bool:
    """Return True if the Program Level cell contains 'WM' (covers OP WM too)."""
    return "WM" in str(cell_value or "").upper()


def is_anthem_payer(payer: str) -> bool:
    """Return True if the payer contains 'anthem'."""
    return "anthem" in (payer or "").lower()


def is_bcb_anthem_ct_php_res_detox(payer: str, service: str) -> bool:
    """Return True if payer is BCB Anthem CT and service is PHP, Residential, or Detox."""
    return ("bcb anthem ct" in (payer or "").lower() and
            any(s in (service or "").lower() for s in ["partial hospitalization", "residential", "detox"]))


class ReportExclusions:
    """The per-run options that keep rows out of the individual staff reports.

    Every one of these is off by default. Excluded rows are still assigned
    an owner in the Masters workbook; they are only left out of the
    individual workbooks. app.py and "Unbilled Step 1.py" both filter their
    reports through excludes(), and assign_staff uses the same check to
    leave excluded rows out of the Jasmine/Cathy half-and-half count, so
    the two reports come out even after the exclusions are applied.

    - exclude_optum: Optum drug screen (utox) rows.
    - exclude_bcb_anthem_ct: BCB Anthem CT PHP/Residential/Detox rows.
    - exclude_anthem_cathy_jasmine: Anthem rows, from Cathy's and
      Jasmine's reports only.
    - exclude_aetna: Aetna rows.
    - exclude_detox_residential: Detox/Residential rows.
    - division_payer_terms/division_terms: rows whose Payer matches one of
      division_payer_terms AND whose GROUPFLD1 matches one of
      division_terms (see payer_excluded_by_division).
    - payer_terms/service_terms: free-text payer/service exclusions,
      limited to the staff named in scope (empty scope = everyone).
    """

    def __init__(self, exclude_optum: bool = False, exclude_bcb_anthem_ct: bool = False,
                 exclude_anthem_cathy_jasmine: bool = False, exclude_aetna: bool = False,
                 exclude_detox_residential: bool = False,
                 payer_terms: list = None, service_terms: list = None, scope: list = None,
                 division_payer_terms: list = None, division_terms: list = None):
        self.exclude_optum = exclude_optum
        self.exclude_bcb_anthem_ct = exclude_bcb_anthem_ct
        self.exclude_anthem_cathy_jasmine = exclude_anthem_cathy_jasmine
        self.exclude_aetna = exclude_aetna
        self.exclude_detox_residential = exclude_detox_residential
        self.payer_terms = payer_terms or []
        self.service_terms = service_terms or []
        self.scope = scope or []
        self.division_payer_terms = division_payer_terms or []
        self.division_terms = division_terms or []

    def excludes(self, staff: str, payer: str, service: str, division: str) -> bool:
        """Return True if this row should be left out of staff's report."""
        payer = payer or ""
        service = service or ""
        if self.exclude_optum and _is_drug_screen(service) and "optum" in payer.lower():
            return True
        if self.exclude_bcb_anthem_ct and is_bcb_anthem_ct_php_res_detox(payer, service):
            return True
        if (self.exclude_anthem_cathy_jasmine and staff in ("Cathy", "Jasmine")
                and is_anthem_payer(payer)):
            return True
        if self.exclude_aetna and is_aetna_payer(payer):
            return True
        if self.exclude_detox_residential and _is_programming_service(service):
            return True
        if payer_excluded_by_division(payer, division or "",
                                      self.division_payer_terms, self.division_terms):
            return True
        if not self.scope or staff in self.scope:
            if matches_any_term(payer, self.payer_terms):
                return True
            if matches_any_term(service, self.service_terms):
                return True
        return False

    def excludes_from_split(self, payer: str, service: str, division: str) -> bool:
        """Return True if the row would be left out of the report whichever of
        Jasmine or Cathy it went to, so it shouldn't count toward the split.

        A free-text exclusion scoped to only one of the two can't be known
        until the split decides who owns the row, so it doesn't count here;
        it is still applied to that person's report afterwards.
        """
        return (self.excludes("Jasmine", payer, service, division)
                and self.excludes("Cathy", payer, service, division))


# Staff names assign_staff already routes rows to on its own. A custom
# report's name (see assign_staff's custom_report_name) must not collide
# with one of these, or its rows would be indistinguishable from that
# staff's own.
RESERVED_STAFF_NAMES = ("Jasmine", "CB", "Melissa", "Cathy", "Unable to Bill")


def validate_custom_report_name(name: str) -> None:
    """Raise ValueError if name collides with a reserved staff name."""
    if name and name.strip().lower() in {n.lower() for n in RESERVED_STAFF_NAMES}:
        raise ValueError(
            f"'{name}' is already a reserved staff name "
            f"({', '.join(RESERVED_STAFF_NAMES)}); choose a different name "
            "for the custom report."
        )


def assign_staff(ws, date_token: str = None, include_programming: bool = False,
                 skip_cathy: bool = False, split_day: str = None,
                 exclusions: ReportExclusions = None,
                 custom_report_name: str = None, custom_report_payer_terms: list = None,
                 custom_report_professional_only: bool = True):
    """Assign staff names based on the standard daily billing rules.

    This is the single copy of the assignment engine: app.py (the Streamlit
    app) and "Unbilled Step 1.py" (the CLI script) both import it from here
    instead of each keeping their own copy, so a rule change can no longer
    happen in one and not the other.

    - Self Pay (GROUPFLD2 == "Self Pay") always goes to CB; every service
      bills every day, with no exceptions.
    - Jasmine and Cathy only ever receive GROUPFLD2 == "Insurance" rows.
    - Every day of the week, everything that goes to Jasmine or Cathy is
      split exactly in half between them. The "shared pool" is every
      Insurance row either would get: Professional rows (Claim Type
      CMS-1500 or UB-04, which bill every day), IOP (including Telemed IOP,
      every day), billable Programming (Detox, Residential) or e-care for
      that weekday, and any other billable Insurance row whose Claim Type
      is not CMS-1500/UB-04 (i.e. institutional/837I). The pool is sorted
      alphabetically by Client and cut at its exact midpoint — even if that
      falls in the middle of a letter, or between two rows for the same
      client. With an odd row count the second half gets the one extra row.
    - Rows the per-run report exclusions (see ReportExclusions) would keep
      out of both Jasmine's and Cathy's reports don't count toward the
      split, so the halves are equal in what actually reaches the two
      reports. They still get an owner in the Masters workbook: whoever's
      half they sort into alphabetically.
    - The halves alternate daily (see split_day_for_date_token): on Day A
      Jasmine gets the first (A-M) half and Cathy the second (N-Z) half;
      on Day B they swap.
    - Any other Insurance row (not billable that day), or any row that is
      neither Self Pay nor Insurance, is Unable to Bill.
    - Melissa (WM/OP WM Program Level, PHP/Partial Hospitalization, or
      Aetna/Humana Detox/Residential) and the O'Flynn Karen OP
      Chappaqua/OP NYC "Unable to Bill" rule take priority over all of the
      above. PHP rows are assigned to Melissa every day in the Masters
      spreadsheet (except O'Flynn Karen OP Chappaqua/OP NYC PHP, which
      stays Unable to Bill); she does not get an individual report, and
      PHP is billed only on Tuesdays as an operational matter.

    Optional, per-run overrides (all off by default) sit on top of the
    schedule above:
    - include_programming: Programming (Detox, Residential) is billable
      regardless of the weekday, so it can be worked on a Monday or
      Wednesday. E-care is unaffected and stays Tuesday-only.
    - skip_cathy: Cathy is given no rows at all; the whole shared pool
      goes to Jasmine.
    - split_day: "A" or "B" to force that day's split for this run instead
      of deriving it from date_token.
    - exclusions: the run's ReportExclusions, so rows excluded from both
      reports are left out of the split count (see above).
    - custom_report_name/custom_report_payer_terms/
      custom_report_professional_only: a custom report for routing a specific payer's rows to a different named staff member
      without a code change. When custom_report_name and
      custom_report_payer_terms are both set, every Insurance row whose
      Payer contains one of those terms is that staff's, ahead of the
      shared pool. custom_report_professional_only (default True)
      restricts it to CMS-1500/UB-04 claim types; set it False to match
      any claim type instead. The name must not collide with a reserved
      staff name (see RESERVED_STAFF_NAMES/validate_custom_report_name).

    Args:
        ws: Worksheet to process
        date_token: Date string in MMDDYYYY format (from filename). If None, uses current date.
        include_programming: When True, bill Programming regardless of weekday.
        skip_cathy: When True, assign Cathy nothing at all; Jasmine takes
            the whole shared pool.
        split_day: "A" or "B" to force the split for this run; None (the
            default) derives it from date_token.
        exclusions: ReportExclusions for this run; None means no
            exclusions.
        custom_report_name: When set (with custom_report_payer_terms), the
            staff name to assign matching rows to.
        custom_report_payer_terms: When set (with custom_report_name), payer
            terms (case-insensitive substring match) that route a row to
            custom_report_name.
        custom_report_professional_only: When True (default), the custom
            report only claims Professional (CMS-1500/UB-04) rows. Set False to match any claim type.

    Returns:
        The split day that was used for this run, "A" or "B".
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
        elif header == "Claim Type":
            cols['claim_type'] = col

    if not all(k in cols for k in ['group', 'service', 'payer', 'claim_type', 'client']):
        raise ValueError(
            "Missing required columns: GROUPFLD2, Service, Payer, Claim Type, or Client"
        )

    weekday, did_fallback = parse_weekday_from_token(date_token)
    if did_fallback:
        print(f"Failed to parse date_token '{date_token}', using current weekday={weekday}")
    else:
        print(f"Using date from filename: {date_token} (weekday={weekday})")

    if split_day is None:
        split_day, _ = split_day_for_date_token(date_token)
    else:
        split_day = str(split_day).strip().upper()
        if split_day not in SPLIT_DAYS:
            raise ValueError(f"split_day must be 'A' or 'B', not {split_day!r}")
    print(f"Jasmine/Cathy split: Day {split_day}")

    if exclusions is None:
        exclusions = ReportExclusions()

    custom_report_active = bool(custom_report_name and custom_report_payer_terms)

    row_data_map = {}
    fixed_staff = {}        # original_row -> staff already decided
    shared_pool = []        # (original_row, client, excluded) to be split in half between Jasmine and Cathy
    other_rows = []         # original_row order for every row not in the shared pool

    for row in range(2, ws.max_row + 1):
        row_data_map[row] = [ws.cell(row, col).value for col in range(1, ws.max_column + 1)]

        group = str(ws.cell(row, cols['group']).value or "").strip()
        service = str(ws.cell(row, cols['service']).value or "")
        payer = str(ws.cell(row, cols['payer']).value or "").lower()
        claim_type = str(ws.cell(row, cols['claim_type']).value or "")

        # CB: Self Pay bills every service every day, no exceptions. This
        # overrides every other rule.
        if group == "Self Pay":
            fixed_staff[row] = "CB"
            other_rows.append(row)
            continue

        staff = None

        # WM / OP WM Program Level → Melissa
        if 'program_level' in cols and is_wm_program_level(ws.cell(row, cols['program_level']).value):
            staff = "Melissa"

        # Unable to Bill: Billing Provider = "O'Flynn, Karen" + GROUPFLD1 = "OP Chappaqua" or "OP NYC"
        if not staff and 'billing_provider' in cols and 'group_fld1' in cols:
            billing_provider = str(ws.cell(row, cols['billing_provider']).value or "").strip()
            group_fld1 = str(ws.cell(row, cols['group_fld1']).value or "").strip()
            if (billing_provider == "O'Flynn, Karen" and
                    group_fld1 in ("OP Chappaqua", "OP NYC")):
                staff = "Unable to Bill"

        # Melissa: (Detox or Residential) + (Aetna or Humana), but not drug screens
        if not staff:
            service_lower = service.lower()
            has_detox_res = ("detox" in service_lower or "residential" in service_lower)
            has_insurance = "aetna" in payer or "humana" in payer
            if has_detox_res and has_insurance and not _is_drug_screen(service):
                staff = "Melissa"

        # PHP/Partial Hospitalization always goes to Melissa, every day
        # (after the O'Flynn Karen rule, whose PHP stays Unable to Bill).
        if not staff and is_php_service(service):
            staff = "Melissa"

        if staff:
            fixed_staff[row] = staff
            other_rows.append(row)
            continue

        # Jasmine and Cathy only ever receive Insurance rows.
        if group != "Insurance":
            fixed_staff[row] = "Unable to Bill"
            other_rows.append(row)
            continue

        # Custom report (optional): routes a specific payer's rows to a
        # different staff member, ahead of the shared pool.
        if (custom_report_active
                and (not custom_report_professional_only or is_professional_claim_type(claim_type))
                and matches_any_term(payer, custom_report_payer_terms)):
            fixed_staff[row] = custom_report_name
            other_rows.append(row)
            continue

        # Everything else billable is Jasmine's and Cathy's to share: IOP
        # and Professional rows bill every day; Programming/e-care/other
        # institutional rows only when the weekday schedule allows.
        if (not is_iop_service(service)
                and not is_professional_claim_type(claim_type)
                and is_non_billable_service_for_weekday(
                    service, weekday, include_programming=include_programming)):
            fixed_staff[row] = "Unable to Bill"
            other_rows.append(row)
            continue

        client = str(ws.cell(row, cols['client']).value or "").strip()
        division = (str(ws.cell(row, cols['group_fld1']).value or "")
                    if 'group_fld1' in cols else "")
        excluded = exclusions.excludes_from_split(
            str(ws.cell(row, cols['payer']).value or ""), service, division)
        shared_pool.append((row, client, excluded))

    # Move the shared pool to the top of the sheet, sorted alphabetically by
    # Client; every other row keeps its original relative order after that.
    shared_pool.sort(key=lambda x: x[1].lower())

    # Cut the sorted pool at the exact midpoint of the rows that will
    # actually reach a report (excluded rows don't count), even if that
    # lands mid-letter or mid-client. Day A: Jasmine takes the first (A-M)
    # half, Cathy the second (N-Z) half; Day B: the reverse. With
    # skip_cathy, Jasmine takes both halves.
    midpoint = sum(1 for _, _, excluded in shared_pool if not excluded) // 2
    if split_day == "A":
        first_half_staff, second_half_staff = "Jasmine", "Cathy"
    else:
        first_half_staff, second_half_staff = "Cathy", "Jasmine"
    if skip_cathy:
        first_half_staff = second_half_staff = "Jasmine"

    ordered_rows = [row for row, _, _ in shared_pool] + other_rows
    new_row_pos = 2
    for original_row in ordered_rows:
        for col in range(1, ws.max_column + 1):
            ws.cell(new_row_pos, col).value = row_data_map[original_row][col - 1]
        new_row_pos += 1

    # An excluded row still gets an owner in the Masters workbook: whoever's
    # half it sorts into, i.e. the first half until `midpoint` counted rows
    # have been handed out.
    new_row_pos = 2
    counted = 0
    for _, _, excluded in shared_pool:
        in_first_half = counted < midpoint
        ws.cell(new_row_pos, 1).value = first_half_staff if in_first_half else second_half_staff
        if not excluded:
            counted += 1
        new_row_pos += 1
    for original_row in other_rows:
        ws.cell(new_row_pos, 1).value = fixed_staff[original_row]
        new_row_pos += 1

    print("Staff assignment complete")
    return split_day


def finalize_workbook(wb, include_batch_billings: bool = False, include_iop_status: bool = False,
                      skip_status_columns: bool = False):
    """Add Status/Comments columns and validation for Jasmine/Cathy/custom exports.

    Args:
        wb: Workbook to finalize.
        include_batch_billings: When True, add 'Batch Billings' to the dropdown
            (Jasmine, Cathy, and custom reports).
        include_iop_status: When True, add 'IOP' to the dropdown (Jasmine,
            Cathy, and custom reports).
        skip_status_columns: When True, skip adding Status/Comments columns (for CB/self-pay).
    """
    ws = wb.active

    # Bold header row
    for col in range(1, ws.max_column + 1):
        ws.cell(1, col).font = openpyxl.styles.Font(bold=True)

    # Insert two columns at E (unless skipped for CB/self-pay)
    if not skip_status_columns:
        ws.insert_cols(5, 2)
        ws.cell(1, 5).value = "Status"
        ws.cell(1, 6).value = "Comments"
        ws.cell(1, 5).font = openpyxl.styles.Font(bold=True)
        ws.cell(1, 6).font = openpyxl.styles.Font(bold=True)

        # Create Sheet2 with validation list
        ws_list = wb.create_sheet("Sheet2")
        status_items = ["Billed", "Unable to Bill", "Contractual Adj", "Incomplete Billings",
                        "Utox Batch", "Inclusive Services"]
        if include_batch_billings:
            status_items.append("Batch Billings")
        if include_iop_status:
            status_items.append("IOP")

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


def is_non_billable_service_for_weekday(
    service: str,
    weekday: int,
    is_professional: bool = False,
    self_pay: bool = False,
    include_programming: bool = False,
) -> bool:
    """
    Return True if the given service (string) should be treated as non-billable
    on the specified weekday.

    - service: original service string (case-insensitive checks will be used)
    - weekday: 0=Monday .. 6=Sunday
    - is_professional: True when the row's Claim Type is CMS-1500 or UB-04.
      Professional rows bill every day of the week.
    - self_pay: when True, every service is billable every day, with no
      exceptions (including e-care).
    - include_programming: one-off override. When True, Programming (Detox,
      Residential) is billable regardless of the weekday, so it can be
      included on a Monday or Wednesday. It does not affect e-care, which
      stays Tuesday-only.
    """
    s = (service or "").lower().strip()

    # Self Pay bills every service every day, no exceptions.
    if self_pay:
        return False

    # Professional (CMS-1500) services bill every day of the week.
    if is_professional:
        return False

    # e-care is only billed on Tuesdays.
    if _is_ecare(s):
        return weekday != 1

    # IOP (including Telemed IOP) bills every day of the week.
    if is_iop_service(s):
        return False

    # Programming (Detox, Residential) bills Tue/Thu/Fri + weekends, unless
    # the one-off include_programming override is on for this run.
    if _is_programming_service(s):
        if include_programming:
            return False
        return weekday not in (1, 3, 4, 5, 6)

    # Anything else is neither Professional nor Programming nor e-care, and
    # isn't addressed by the daily schedule, so it is not billable.
    return True
