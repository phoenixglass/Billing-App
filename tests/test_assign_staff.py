"""
End-to-end tests for assign_staff: the daily half-and-half Jasmine/Cathy
split and its Day A/Day B rotation, report exclusions being applied before
the split, "give Cathy nothing", Insurance PHP going to Melissa, the custom
report, and the include_programming override, exercised against a real
worksheet.

assign_staff lives in billing_rules.py; app.py and "Unbilled Step 1.py"
both import it from there. app.py also imports Streamlit, so these tests
load the standalone script instead, which re-exports the same function.

Run: python tests/test_assign_staff.py
Requires: openpyxl
"""
import importlib.util
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).parent.parent
sys.path.insert(0, str(REPO_ROOT))

from openpyxl import Workbook


def _load_unbilled_module():
    """Import "Unbilled Step 1.py" (its filename is not a valid module name)."""
    spec = importlib.util.spec_from_file_location(
        "unbilled_step_1", REPO_ROOT / "Unbilled Step 1.py"
    )
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


unbilled = _load_unbilled_module()
assign_staff = unbilled.assign_staff
ReportExclusions = unbilled.ReportExclusions

HEADERS = ["Staff/Status", "GROUPFLD1", "GROUPFLD2", "Service", "Payer",
           "Billing Provider", "Program Level", "Client", "Claim Type"]

# MMDDYYYY tokens for the two days most of these tests use. 09/23/2026 is
# the Day A anchor for the Jasmine/Cathy split, so 09/01/2026 (22 days
# earlier) is Day A and 09/02/2026 (21 days earlier) is Day B.
WEDNESDAY = "09022026"  # Day B
TUESDAY = "09012026"    # Day A


def _sheet(rows):
    """Build a worksheet from (group, service, payer, client, claim_type) rows."""
    wb = Workbook()
    ws = wb.active
    ws.append(HEADERS)
    for group, service, payer, client, claim_type in rows:
        ws.append(["", "OP Westchester", group, service, payer, "Smith, John",
                   "OP", client, claim_type])
    return ws


def _staff_by_client(ws):
    """Map Client -> assigned staff after assign_staff has reordered the rows.

    Column 1 is the Staff/Status column, so this reads the same assignment
    the Masters workbook carries for each row.
    """
    client_col = HEADERS.index("Client") + 1
    return {
        ws.cell(row, client_col).value: ws.cell(row, 1).value
        for row in range(2, ws.max_row + 1)
    }


def _mixed_payer_rows():
    return [
        ("Insurance", "Individual Therapy", "Oxford", "Adams, Ann", "CMS-1500"),
        ("Insurance", "Individual Therapy", "ConnectiCare", "Baker, Bob", "UB-04"),
        ("Insurance", "Group Therapy", "UBH", "Carter, Cal", "CMS-1500"),
        # Not Professional, and not billable on a Wednesday.
        ("Insurance", "Individual Therapy", "Oxford", "Diaz, Dee", "837I"),
        ("Insurance", "Individual Therapy", "Aetna", "Evans, Eve", "CMS-1500"),
        ("Insurance", "Individual Therapy", "Optum", "Frank, Fay", "CMS-1500"),
    ]


def _assignments(ws):
    """Column 1 (Staff/Status) for every data row, in sheet order."""
    return [ws.cell(row, 1).value for row in range(2, ws.max_row + 1)]


def _pool_rows(n, prefix="Pool", payer="Magellan", service="Individual Therapy"):
    return [("Insurance", service, payer, f"{prefix}{i:04d}", "CMS-1500")
            for i in range(n)]


def test_day_a_jasmine_gets_a_to_m_and_cathy_gets_n_to_z():
    """Day A: Jasmine takes the first (A-M) half of the sorted pool, Cathy the second."""
    clients = ["Zeller, Zoe", "Adams, Ann", "Nolan, Ned", "Miller, Mia",
               "Baker, Bob", "Young, Yan"]
    rows = [("Insurance", "Individual Therapy", "Magellan", c, "CMS-1500")
            for c in clients]
    ws = _sheet(rows)
    assert assign_staff(ws, "09232026") == "A"
    staff = _staff_by_client(ws)

    for client in ("Adams, Ann", "Baker, Bob", "Miller, Mia"):
        assert staff[client] == "Jasmine", client
    for client in ("Nolan, Ned", "Young, Yan", "Zeller, Zoe"):
        assert staff[client] == "Cathy", client


def test_day_b_swaps_the_halves():
    """Day B (the next day): Jasmine takes N-Z and Cathy takes A-M."""
    clients = ["Zeller, Zoe", "Adams, Ann", "Nolan, Ned", "Miller, Mia",
               "Baker, Bob", "Young, Yan"]
    rows = [("Insurance", "Individual Therapy", "Magellan", c, "CMS-1500")
            for c in clients]
    ws = _sheet(rows)
    assert assign_staff(ws, "09242026") == "B"
    staff = _staff_by_client(ws)

    for client in ("Adams, Ann", "Baker, Bob", "Miller, Mia"):
        assert staff[client] == "Cathy", client
    for client in ("Nolan, Ned", "Young, Yan", "Zeller, Zoe"):
        assert staff[client] == "Jasmine", client


def test_split_is_an_exact_half_even_when_letters_are_lopsided():
    """The cut is at the pool's midpoint, not at the letter M, so the halves
    stay equal even when most clients fall in A-M."""
    rows = _pool_rows(8, prefix="Adams")
    rows += _pool_rows(2, prefix="Young")
    ws = _sheet(rows)
    assign_staff(ws, "09232026")  # Day A
    assignments = _assignments(ws)

    assert assignments.count("Jasmine") == 5
    assert assignments.count("Cathy") == 5
    # The sheet is sorted by Client, so Jasmine's half is the top of it.
    assert assignments[:5] == ["Jasmine"] * 5


def test_odd_pool_gives_the_extra_row_to_the_second_half():
    rows = _pool_rows(7)
    ws = _sheet(rows)
    assign_staff(ws, "09232026")  # Day A: second (N-Z) half is Cathy's
    assignments = _assignments(ws)
    assert assignments.count("Jasmine") == 3
    assert assignments.count("Cathy") == 4


def test_split_applies_on_weekends_too():
    """Cathy gets her half on a Saturday and Sunday as well."""
    rows = _pool_rows(50)
    for token in ("09262026", "09272026"):  # Saturday, Sunday
        ws = _sheet(rows)
        assign_staff(ws, token)
        assignments = _assignments(ws)
        assert assignments.count("Cathy") == 25, token
        assert assignments.count("Jasmine") == 25, token


def test_split_day_can_be_forced_for_a_run():
    rows = [("Insurance", "Individual Therapy", "Magellan", "Adams, Ann", "CMS-1500"),
            ("Insurance", "Individual Therapy", "Magellan", "Zeller, Zoe", "CMS-1500")]

    ws = _sheet(rows)
    # 09/23/2026 is Day A on its own; force Day B instead.
    assert assign_staff(ws, "09232026", split_day="b") == "B"
    staff = _staff_by_client(ws)
    assert staff["Adams, Ann"] == "Cathy"
    assert staff["Zeller, Zoe"] == "Jasmine"

    ws = _sheet(rows)
    try:
        assign_staff(ws, "09232026", split_day="C")
        assert False, "expected ValueError for split_day='C'"
    except ValueError:
        pass


def test_split_covers_iop_programming_ecare_and_institutional_rows():
    """Everything Jasmine or Cathy would get is in the shared pool, not just
    Professional rows."""
    rows = [
        ("Insurance", "Telemed IOP", "Magellan", "Adams, Ann", "CMS-1500"),
        ("Insurance", "IOP", "Magellan", "Baker, Bob", "837I"),
        ("Insurance", "Detox Admission", "Magellan", "Carter, Cal", "837I"),
        ("Insurance", "Residential Program", "Magellan", "Nolan, Ned", "837I"),
        ("Insurance", "E-Care Individual", "Magellan", "Owens, Olu", "837I"),
        ("Insurance", "Individual Therapy", "Magellan", "Price, Pam", "CMS-1500"),
    ]
    ws = _sheet(rows)
    assign_staff(ws, TUESDAY)  # 09/01/2026: Tuesday, Day A
    staff = _staff_by_client(ws)

    for client in ("Adams, Ann", "Baker, Bob", "Carter, Cal"):
        assert staff[client] == "Jasmine", client
    for client in ("Nolan, Ned", "Owens, Olu", "Price, Pam"):
        assert staff[client] == "Cathy", client


def test_non_billable_rows_stay_out_of_the_split():
    """Wednesday Detox/e-care (institutional) are Unable to Bill, not split."""
    rows = [
        ("Insurance", "Detox Admission", "Magellan", "Detox, Dan", "837I"),
        ("Insurance", "E-Care Individual", "Magellan", "Ecare, Ed", "837I"),
        ("Insurance", "Individual Therapy", "Magellan", "Adams, Ann", "CMS-1500"),
        ("Insurance", "Individual Therapy", "Magellan", "Zeller, Zoe", "CMS-1500"),
    ]
    ws = _sheet(rows)
    assign_staff(ws, WEDNESDAY)  # 09/02/2026: Day B
    staff = _staff_by_client(ws)
    assert staff["Detox, Dan"] == "Unable to Bill"
    assert staff["Ecare, Ed"] == "Unable to Bill"
    assert staff["Adams, Ann"] == "Cathy"
    assert staff["Zeller, Zoe"] == "Jasmine"


def _report_counts(ws, exclusions):
    """How many rows would reach Jasmine's and Cathy's workbooks — the same
    filter app.py and the CLI apply when they build the individual reports."""
    payer_col = HEADERS.index("Payer") + 1
    service_col = HEADERS.index("Service") + 1
    division_col = HEADERS.index("GROUPFLD1") + 1
    counts = {"Jasmine": 0, "Cathy": 0}
    for row in range(2, ws.max_row + 1):
        staff = ws.cell(row, 1).value
        if staff not in counts:
            continue
        if exclusions.excludes(staff, ws.cell(row, payer_col).value,
                               ws.cell(row, service_col).value,
                               ws.cell(row, division_col).value):
            continue
        counts[staff] += 1
    return counts


def test_exclusions_come_before_the_split():
    """Excluded rows don't count toward the half-and-half split, so the two
    reports are still even after the exclusion — even when every excluded
    row sorts into the same half."""
    rows = _pool_rows(10, prefix="Adams", payer="Aetna")   # all A-M side
    rows += _pool_rows(10, prefix="Baker")
    rows += _pool_rows(10, prefix="Young")
    exclusions = ReportExclusions(exclude_aetna=True)

    ws = _sheet(rows)
    assign_staff(ws, "09232026", exclusions=exclusions)  # Day A
    assert _report_counts(ws, exclusions) == {"Jasmine": 10, "Cathy": 10}

    # Without passing the exclusions to the split, the Aetna rows would have
    # eaten Jasmine's half, leaving her 5 reportable rows to Cathy's 15.
    ws = _sheet(rows)
    assign_staff(ws, "09232026")
    assert _report_counts(ws, exclusions) == {"Jasmine": 5, "Cathy": 15}


def test_excluded_rows_still_get_an_owner_in_masters():
    """Excluded rows stay in the Masters workbook, owned by whoever's half
    they sort into."""
    rows = [
        ("Insurance", "Individual Therapy", "Aetna", "Adams, Ann", "CMS-1500"),
        ("Insurance", "Individual Therapy", "Magellan", "Baker, Bob", "CMS-1500"),
        ("Insurance", "Individual Therapy", "Magellan", "Carter, Cal", "CMS-1500"),
        ("Insurance", "Individual Therapy", "Aetna", "Nolan, Ned", "CMS-1500"),
        ("Insurance", "Individual Therapy", "Magellan", "Young, Yan", "CMS-1500"),
        ("Insurance", "Individual Therapy", "Magellan", "Zeller, Zoe", "CMS-1500"),
    ]
    ws = _sheet(rows)
    assign_staff(ws, "09232026", exclusions=ReportExclusions(exclude_aetna=True))
    staff = _staff_by_client(ws)

    # Counted rows: Baker, Carter | Young, Zeller.
    assert staff["Baker, Bob"] == "Jasmine"
    assert staff["Carter, Cal"] == "Jasmine"
    assert staff["Young, Yan"] == "Cathy"
    assert staff["Zeller, Zoe"] == "Cathy"
    # Adams sorts into Jasmine's half, Nolan (after the cut) into Cathy's.
    assert staff["Adams, Ann"] == "Jasmine"
    assert staff["Nolan, Ned"] == "Cathy"
    assert len(_assignments(ws)) == 6


def test_every_report_exclusion_is_applied_before_the_split():
    """Each exclusion that covers both Jasmine and Cathy is left out of the
    split count."""
    base = _pool_rows(10, prefix="Young")
    cases = [
        (ReportExclusions(exclude_aetna=True), "Aetna", "Individual Therapy"),
        (ReportExclusions(exclude_anthem_cathy_jasmine=True), "Anthem BCBS", "Individual Therapy"),
        (ReportExclusions(exclude_optum=True), "Optum", "Drug Screen"),
        (ReportExclusions(exclude_bcb_anthem_ct=True), "BCB Anthem CT", "Residential Program"),
        (ReportExclusions(exclude_detox_residential=True), "Magellan", "Detox Admission"),
        (ReportExclusions(payer_terms=["cigna"]), "Cigna", "Individual Therapy"),
        (ReportExclusions(service_terms=["group"]), "Magellan", "Group Therapy"),
        (ReportExclusions(payer_terms=["cigna"], scope=["Jasmine", "Cathy"]),
         "Cigna", "Individual Therapy"),
        (ReportExclusions(division_payer_terms=["magellan"],
                          division_terms=["westchester"]), "Magellan", "Individual Therapy"),
    ]
    for exclusions, payer, service in cases:
        rows = [("Insurance", service, payer, f"Adams{i:04d}", "CMS-1500")
                for i in range(6)]
        # The division case excludes every Magellan row in OP Westchester,
        # base rows included, so give the base rows a payer it doesn't match.
        rows += [(g, s, "Beacon", c, ct) for g, s, _, c, ct in base]
        ws = _sheet(rows)
        # Saturday, so Detox/Residential are billable and in the pool.
        assign_staff(ws, "09262026", exclusions=exclusions)
        assert _report_counts(ws, exclusions) == {"Jasmine": 5, "Cathy": 5}, (payer, service)


def test_exclusion_scoped_to_one_person_sends_those_rows_to_the_other():
    """A free-text exclusion limited to only Jasmine sends the matching rows
    to Cathy instead, and the rest are split around them so the two
    reports still come out even — evenness beats the alphabet."""
    rows = _pool_rows(4, prefix="Adams", payer="Cigna")
    rows += _pool_rows(4, prefix="Baker")
    rows += _pool_rows(4, prefix="Young")
    exclusions = ReportExclusions(payer_terms=["cigna"], scope=["Jasmine"])

    ws = _sheet(rows)
    assign_staff(ws, "09232026", exclusions=exclusions)  # Day A: Jasmine A-M
    staff = _staff_by_client(ws)

    # The Cigna rows would be dropped from Jasmine's report, so they're Cathy's.
    for i in range(4):
        assert staff[f"Adams{i:04d}"] == "Cathy"
    # That puts Cathy at 4 already, so Jasmine takes the next 6 free rows
    # alphabetically and Cathy the last 2: 6 each.
    assert [staff[f"Baker{i:04d}"] for i in range(4)] == ["Jasmine"] * 4
    assert [staff[f"Young{i:04d}"] for i in range(4)] == ["Jasmine", "Jasmine", "Cathy", "Cathy"]
    assert _report_counts(ws, exclusions) == {"Jasmine": 6, "Cathy": 6}


def test_scoped_exclusion_works_on_day_b_and_for_cathy_too():
    """Same balancing when the exclusion is scoped to Cathy, on a Day B."""
    rows = _pool_rows(3, prefix="Young", payer="Cigna")
    rows += _pool_rows(7, prefix="Adams")
    exclusions = ReportExclusions(payer_terms=["cigna"], scope=["Cathy"])

    ws = _sheet(rows)
    assign_staff(ws, "09242026", exclusions=exclusions)  # Day B: Cathy A-M
    staff = _staff_by_client(ws)
    for i in range(3):
        assert staff[f"Young{i:04d}"] == "Jasmine"
    assert _report_counts(ws, exclusions) == {"Jasmine": 5, "Cathy": 5}


def test_scoped_exclusion_too_big_to_balance_gets_as_close_as_possible():
    """If the one-sided rows alone are more than half, the other person
    takes every remaining row."""
    rows = _pool_rows(8, prefix="Adams", payer="Cigna")
    rows += _pool_rows(2, prefix="Young")
    exclusions = ReportExclusions(payer_terms=["cigna"], scope=["Jasmine"])

    ws = _sheet(rows)
    assign_staff(ws, "09232026", exclusions=exclusions)
    assert _report_counts(ws, exclusions) == {"Jasmine": 2, "Cathy": 8}
    assert all(value for value in _assignments(ws))


def test_scoped_and_blanket_exclusions_together_stay_even():
    rows = _pool_rows(4, prefix="Adams", payer="Aetna")    # excluded for both
    rows += _pool_rows(3, prefix="Carter", payer="Cigna")  # excluded for Cathy
    rows += _pool_rows(9, prefix="Nolan")
    exclusions = ReportExclusions(exclude_aetna=True, payer_terms=["cigna"],
                                  scope=["Cathy"])
    for token in ("09232026", "09242026"):
        ws = _sheet(rows)
        assign_staff(ws, token, exclusions=exclusions)
        assert _report_counts(ws, exclusions) == {"Jasmine": 6, "Cathy": 6}, token
        assert len(_assignments(ws)) == 16


def test_skip_cathy_ignores_scoped_exclusions_for_ownership():
    """With Cathy skipped, Jasmine still owns every pool row."""
    rows = _pool_rows(4, prefix="Adams", payer="Cigna")
    rows += _pool_rows(4, prefix="Young")
    ws = _sheet(rows)
    assign_staff(ws, "09232026", skip_cathy=True,
                 exclusions=ReportExclusions(payer_terms=["cigna"], scope=["Jasmine"]))
    assert set(_assignments(ws)) == {"Jasmine"}


def test_former_cathy_payers_are_just_split():
    """Cathy no longer gets specific payers: Oxford/ConnectiCare/UBH (and her
    old wider list) are split with Jasmine like everything else, IOP
    included."""
    ws = _sheet([
        ("Insurance", "Individual Therapy", "Oxford", "Adams, Ann", "CMS-1500"),
        ("Insurance", "IOP", "ConnectiCare", "Baker, Bob", "UB-04"),
        ("Insurance", "Individual Therapy", "Emblem (Optum)", "Nolan, Ned", "CMS-1500"),
        ("Insurance", "Individual Therapy", "UBH", "Young, Yan", "CMS-1500"),
    ])
    assign_staff(ws, "09232026")  # Day A
    staff = _staff_by_client(ws)
    assert staff["Adams, Ann"] == "Jasmine"
    assert staff["Baker, Bob"] == "Jasmine"
    assert staff["Nolan, Ned"] == "Cathy"
    assert staff["Young, Yan"] == "Cathy"


def test_higher_priority_rules_still_win_over_the_split():
    """Self Pay, WM, and PHP keep their owners; only the rest is split."""
    ws = _sheet([
        ("Self Pay", "Individual Therapy", "Oxford", "Self, Sam", "CMS-1500"),
        ("Insurance", "Individual Therapy", "Oxford", "Wm, Wes", "CMS-1500"),
        ("Insurance", "Partial Hospitalization", "Oxford", "Php, Pat", "CMS-1500"),
        ("Insurance", "Individual Therapy", "UBH", "Cathy, Cam", "CMS-1500"),
    ])
    # WM is decided by the Program Level column, so set it on Wes's row.
    program_level_col = HEADERS.index("Program Level") + 1
    ws.cell(3, program_level_col).value = "OP WM"

    assign_staff(ws, WEDNESDAY)  # Day B; Cam is the whole pool (second half)
    staff = _staff_by_client(ws)

    assert staff["Self, Sam"] == "CB"        # Self Pay is always CB's
    assert staff["Wm, Wes"] == "Melissa"     # only Melissa bills WM
    assert staff["Php, Pat"] == "Melissa"    # PHP is Melissa's
    assert staff["Cathy, Cam"] == "Jasmine"


def test_insurance_php_goes_to_melissa_but_unable_to_bill_php_stays():
    """Insurance PHP goes to Melissa, whatever the payer, claim type or day —
    except PHP the O'Flynn Karen rule marks Unable to Bill, which stays
    Unable to Bill."""
    rows = [
        ["", "OP NYC", "Insurance", "PHP", "Magellan", "O'Flynn, Karen",
         "OP", "Php, Oflynn", "CMS-1500"],
        ["", "OP Chappaqua", "Insurance", "Partial Hospitalization", "Magellan",
         "O'Flynn, Karen", "OP", "Php, Chappaqua", "837I"],
        ["", "OP Westchester", "Insurance", "Partial Hospitalization",
         "Magellan", "Smith, John", "OP", "Php, Inst", "837I"],
        ["", "OP Westchester", "Insurance", "PHP Day", "Oxford",
         "Smith, John", "OP", "Php, Prof", "CMS-1500"],
    ]
    for token in ("09232026", "09242026"):  # Day A and Day B (Wed, Thu)
        for kwargs in ({}, {"skip_cathy": True}):
            wb = Workbook()
            ws = wb.active
            ws.append(HEADERS)
            for row in rows:
                ws.append(list(row))
            assign_staff(ws, token, **kwargs)
            staff = _staff_by_client(ws)
            assert staff["Php, Oflynn"] == "Unable to Bill", (token, kwargs)
            assert staff["Php, Chappaqua"] == "Unable to Bill", (token, kwargs)
            assert staff["Php, Inst"] == "Melissa", (token, kwargs)
            assert staff["Php, Prof"] == "Melissa", (token, kwargs)


def test_skip_cathy_gives_the_whole_pool_to_jasmine():
    """With skip_cathy, Cathy gets nothing and Jasmine takes the whole pool."""
    rows = _pool_rows(200)
    rows.append(("Self Pay", "Individual Therapy", "Self Pay", "Self, Sam", "CMS-1500"))

    # Without the option, the pool is split in half.
    ws = _sheet(rows)
    assign_staff(ws, WEDNESDAY)
    assignments = _assignments(ws)
    assert assignments.count("Cathy") == 100
    assert assignments.count("Jasmine") == 100

    for token in ("09232026", "09242026"):  # both Day A and Day B
        ws = _sheet(rows)
        assign_staff(ws, token, skip_cathy=True)
        assignments = _assignments(ws)

        assert "Cathy" not in assignments
        assert assignments.count("Jasmine") == 200
        # Nothing is left unassigned, and Self Pay still goes to CB.
        assert assignments.count("CB") == 1
        assert len(assignments) == 201
        assert all(value for value in assignments)


def test_programming_included_on_a_wednesday():
    """include_programming makes Wednesday Detox/Residential billable, and
    therefore part of the Jasmine/Cathy split."""
    rows = [
        ("Insurance", "Detox Admission", "Oxford", "Detox, Dan", "837I"),
        ("Insurance", "Residential Program", "Optum", "Res, Rita", "837I"),
        ("Insurance", "E-Care Individual", "Optum", "Ecare, Ed", "837I"),
    ]

    ws = _sheet(rows)
    assign_staff(ws, WEDNESDAY)
    staff = _staff_by_client(ws)
    assert staff["Detox, Dan"] == "Unable to Bill"
    assert staff["Res, Rita"] == "Unable to Bill"

    ws = _sheet(rows)
    assign_staff(ws, WEDNESDAY, include_programming=True)  # Day B
    staff = _staff_by_client(ws)
    assert staff["Detox, Dan"] == "Cathy"
    assert staff["Res, Rita"] == "Jasmine"
    # e-care is untouched by the override: still Tuesday-only.
    assert staff["Ecare, Ed"] == "Unable to Bill"

    ws = _sheet(rows)
    assign_staff(ws, TUESDAY, include_programming=True)
    staff = _staff_by_client(ws)
    assert staff["Ecare, Ed"] in ("Jasmine", "Cathy")


def test_aetna_programming_still_goes_to_melissa():
    """Including Programming does not move Aetna/Humana Detox off Melissa."""
    ws = _sheet([
        ("Insurance", "Detox Admission", "Aetna", "Aetna, Amy", "837I"),
        ("Insurance", "Residential Program", "Humana", "Humana, Hal", "837I"),
    ])
    assign_staff(ws, WEDNESDAY, include_programming=True)
    staff = _staff_by_client(ws)

    assert staff["Aetna, Amy"] == "Melissa"
    assert staff["Humana, Hal"] == "Melissa"


def _dropdown_options(**kwargs):
    """Run finalize_workbook on a small workbook and read back its Status list."""
    wb = Workbook()
    ws = wb.active
    ws.append(HEADERS)
    ws.append(["Cathy", "OP Westchester", "Insurance", "Individual Therapy",
               "Oxford", "Smith, John", "OP", "Adams, Ann", "CMS-1500"])
    unbilled.finalize_workbook(wb, **kwargs)
    sheet2 = wb["Sheet2"]
    return [sheet2.cell(row, 1).value for row in range(1, sheet2.max_row + 1)]


def test_cathy_status_dropdown_matches_jasmine():
    """Cathy's Status dropdown carries the same options as Jasmine's."""
    base = ["Billed", "Unable to Bill", "Contractual Adj", "Incomplete Billings",
            "Utox Batch", "Inclusive Services"]

    # The plain list: the six shared options (used for CB, which skips
    # Status/Comments entirely, but this checks finalize_workbook's default).
    assert _dropdown_options() == base

    # Jasmine's and Cathy's list: the same six plus Batch Billings and IOP.
    jasmine = _dropdown_options(include_batch_billings=True, include_iop_status=True)
    assert jasmine == base + ["Batch Billings", "IOP"]


def test_custom_report_routes_matching_professional_payer_rows():
    """A custom report claims Professional rows for its payers."""
    ws = _sheet(_mixed_payer_rows())
    assign_staff(ws, WEDNESDAY, custom_report_name="Karen",
                 custom_report_payer_terms=["aetna"])
    staff = _staff_by_client(ws)

    assert staff["Evans, Eve"] == "Karen"  # Aetna, CMS-1500
    # Optum wasn't in the custom report's payer list, so it stays in the
    # Jasmine/Cathy pool (Adams, Baker | Carter, Frank on Day B).
    assert staff["Frank, Fay"] == "Jasmine"
    assert staff["Adams, Ann"] == "Cathy"
    # Oxford isn't in this custom report's payer list ("aetna"), so this row
    # is untouched by it either way; it's a non-Professional claim type
    # outside the daily schedule, so it's Unable to Bill regardless.
    assert staff["Diaz, Dee"] == "Unable to Bill"


def test_custom_report_any_claim_type_when_professional_only_is_false():
    """custom_report_professional_only=False matches any claim type, not just Professional."""
    rows = [
        ("Insurance", "Individual Therapy", "Cigna", "Adams, Ann", "CMS-1500"),
        ("Insurance", "Individual Therapy", "Cigna", "Baker, Bob", "837I"),
    ]
    ws = _sheet(rows)
    assign_staff(ws, WEDNESDAY, custom_report_name="Karen",
                 custom_report_payer_terms=["cigna"],
                 custom_report_professional_only=False)
    staff = _staff_by_client(ws)

    assert staff["Adams, Ann"] == "Karen"
    assert staff["Baker, Bob"] == "Karen"


def test_custom_report_leaves_pool_for_cathy_and_jasmine():
    """Custom report rows leave the shared pool entirely, before the split."""
    rows = _pool_rows(40, prefix="Karen", payer="Cigna")
    rows += _pool_rows(160, payer="Optum")

    ws = _sheet(rows)
    assign_staff(ws, WEDNESDAY, custom_report_name="Karen",
                 custom_report_payer_terms=["cigna"])
    assignments = _assignments(ws)

    assert assignments.count("Karen") == 40
    assert assignments.count("Cathy") == 80
    assert assignments.count("Jasmine") == 80


def test_custom_report_inactive_without_both_name_and_payers():
    """Setting only the name or only the payer terms leaves the standard schedule in place."""
    rows = [("Insurance", "Individual Therapy", "Cigna", "Adams, Ann", "CMS-1500"),
            ("Insurance", "Individual Therapy", "Cigna", "Zeller, Zoe", "CMS-1500")]

    ws = _sheet(rows)
    assign_staff(ws, WEDNESDAY, custom_report_name="Karen")
    assert "Karen" not in _assignments(ws)
    assert sorted(_assignments(ws)) == ["Cathy", "Jasmine"]

    ws = _sheet(rows)
    assign_staff(ws, WEDNESDAY, custom_report_payer_terms=["cigna"])
    assert sorted(_assignments(ws)) == ["Cathy", "Jasmine"]


def test_validate_custom_report_name_rejects_reserved_names():
    """Reserved staff names (case-insensitive) are rejected for the custom report."""
    for name in ("jasmine", "CB", "Melissa", "cathy", "Unable to Bill"):
        try:
            unbilled.validate_custom_report_name(name)
            assert False, f"expected ValueError for reserved name {name!r}"
        except ValueError:
            pass

    # A non-reserved name, and no name at all, are both fine.
    unbilled.validate_custom_report_name("Karen")
    unbilled.validate_custom_report_name("Rosanna")
    unbilled.validate_custom_report_name(None)
    unbilled.validate_custom_report_name("")


if __name__ == '__main__':
    for name, test in sorted(
        (name, obj) for name, obj in list(globals().items())
        if name.startswith('test_') and callable(obj)
    ):
        test()
        print(f"✓ {name} passed")

    print("\nAll tests passed!")
