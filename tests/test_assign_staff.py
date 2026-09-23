"""
End-to-end tests for assign_staff: the daily half-and-half Jasmine/Cathy
split and its Day A/Day B rotation, the Cathy payer carve-out (both payer
lists), "give Cathy nothing", Insurance PHP always going to Melissa, the
custom report (a second, generic Cathy-shaped slot), and the
include_programming override, exercised against a real worksheet.

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


def _cathy_candidate_rows():
    return [
        ("Insurance", "Individual Therapy", "Oxford", "Adams, Ann", "CMS-1500"),
        ("Insurance", "Individual Therapy", "ConnectiCare", "Baker, Bob", "UB-04"),
        ("Insurance", "Group Therapy", "UBH", "Carter, Cal", "CMS-1500"),
        # Not the carve-out's: right payer, but not a Professional claim type.
        ("Insurance", "Individual Therapy", "Oxford", "Diaz, Dee", "837I"),
        # Not the carve-out's: Professional, but a payer it does not cover.
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


def test_cathy_carveout_off_by_default():
    """Without assign_cathy, IOP for her usual payers joins the shared pool
    like any other row instead of going to her outright."""
    ws = _sheet([
        ("Insurance", "IOP", "Oxford", "Adams, Ann", "CMS-1500"),
        ("Insurance", "IOP", "Oxford", "Zeller, Zoe", "CMS-1500"),
    ])
    assign_staff(ws, "09232026")  # Day A
    staff = _staff_by_client(ws)
    assert staff["Adams, Ann"] == "Jasmine"
    assert staff["Zeller, Zoe"] == "Cathy"


def test_cathy_takes_only_professional_rows_for_her_payers():
    """Oxford/ConnectiCare/UBH Professional rows go to Cathy via the carve-out;
    the rest are split with Jasmine as usual."""
    ws = _sheet(_cathy_candidate_rows())
    assign_staff(ws, WEDNESDAY, assign_cathy=True)  # Day B
    staff = _staff_by_client(ws)

    assert staff["Adams, Ann"] == "Cathy"
    assert staff["Baker, Bob"] == "Cathy"
    assert staff["Carter, Cal"] == "Cathy"

    # 837I Oxford is not Professional, so the carve-out does not claim it;
    # on a Wednesday it is not billable either.
    assert staff["Diaz, Dee"] == "Unable to Bill"
    # Evans/Frank are the shared pool: Day B gives the first half to Cathy.
    assert staff["Evans, Eve"] == "Cathy"
    assert staff["Frank, Fay"] == "Jasmine"


def test_cathy_rows_leave_the_shared_pool():
    """A row assigned to Cathy is not also given to Jasmine — no double counting."""
    ws = _sheet(_cathy_candidate_rows())
    assign_staff(ws, WEDNESDAY, assign_cathy=True)

    assignments = _assignments(ws)
    # 3 carve-out rows (Adams/Baker/Carter), then Evans/Frank split one
    # each; Diaz is a non-Professional, non-billable-Wednesday row.
    assert assignments.count("Cathy") == 4
    assert assignments.count("Jasmine") == 1
    assert assignments.count("Unable to Bill") == 1
    # Six rows in, six rows out, each with exactly one owner.
    assert len(assignments) == 6
    assert all(value for value in assignments)


def test_higher_priority_rules_still_win_over_cathy():
    """Self Pay, WM, and PHP keep their owners even for Cathy's payers."""
    ws = _sheet([
        ("Self Pay", "Individual Therapy", "Oxford", "Self, Sam", "CMS-1500"),
        ("Insurance", "Individual Therapy", "Oxford", "Wm, Wes", "CMS-1500"),
        ("Insurance", "Partial Hospitalization", "Oxford", "Php, Pat", "CMS-1500"),
        ("Insurance", "Individual Therapy", "UBH", "Cathy, Cam", "CMS-1500"),
    ])
    # WM is decided by the Program Level column, so set it on Wes's row.
    program_level_col = HEADERS.index("Program Level") + 1
    ws.cell(3, program_level_col).value = "OP WM"

    assign_staff(ws, WEDNESDAY, assign_cathy=True)
    staff = _staff_by_client(ws)

    assert staff["Self, Sam"] == "CB"        # Self Pay is always CB's
    assert staff["Wm, Wes"] == "Melissa"     # only Melissa bills WM
    assert staff["Php, Pat"] == "Melissa"    # PHP is always Melissa's
    assert staff["Cathy, Cam"] == "Cathy"


def test_all_insurance_php_goes_to_melissa():
    """Every Insurance PHP row is Melissa's, whatever the payer, claim type,
    day, or split — including one the O'Flynn Karen rule would otherwise
    mark Unable to Bill."""
    wb = Workbook()
    ws = wb.active
    ws.append(HEADERS)
    ws.append(["", "OP NYC", "Insurance", "PHP", "Magellan", "O'Flynn, Karen",
               "OP", "Php, Oflynn", "CMS-1500"])
    ws.append(["", "OP Westchester", "Insurance", "Partial Hospitalization",
               "Magellan", "Smith, John", "OP", "Php, Inst", "837I"])
    ws.append(["", "OP Westchester", "Insurance", "PHP Day", "Oxford",
               "Smith, John", "OP", "Php, Prof", "CMS-1500"])
    # A non-PHP row from the same O'Flynn Karen/OP NYC combination is still
    # Unable to Bill.
    ws.append(["", "OP NYC", "Insurance", "Individual Therapy", "Magellan",
               "O'Flynn, Karen", "OP", "Oflynn, Other", "CMS-1500"])

    for token in ("09232026", "09242026"):  # Day A and Day B (Wed, Thu)
        for kwargs in ({}, {"assign_cathy": True}, {"skip_cathy": True}):
            wb2 = Workbook()
            ws2 = wb2.active
            for row in ws.iter_rows(values_only=True):
                ws2.append(list(row))
            assign_staff(ws2, token, **kwargs)
            staff = _staff_by_client(ws2)
            assert staff["Php, Oflynn"] == "Melissa", (token, kwargs)
            assert staff["Php, Inst"] == "Melissa", (token, kwargs)
            assert staff["Php, Prof"] == "Melissa", (token, kwargs)
            assert staff["Oflynn, Other"] == "Unable to Bill", (token, kwargs)


def test_cathy_takes_iop_for_her_payers():
    """Every Professional service for Cathy's payers is hers, IOP included."""
    ws = _sheet([
        ("Insurance", "IOP", "Oxford", "Iop, Ida", "CMS-1500"),
        ("Insurance", "Telemed IOP", "ConnectiCare", "Iop, Ivan", "UB-04"),
        ("Insurance", "Detox Admission", "UBH", "Detox, Dora", "CMS-1500"),
        ("Insurance", "E-Care Individual", "Oxford", "Ecare, Ellis", "CMS-1500"),
        # IOP for a payer Cathy does not cover goes to the shared pool.
        ("Insurance", "Telemed IOP", "Optum", "Iop, Otto", "CMS-1500"),
        # IOP for one of her payers, but not a Professional claim type, also
        # goes to the shared pool.
        ("Insurance", "IOP", "Oxford", "Iop, Inst", "837I"),
    ])
    assign_staff(ws, WEDNESDAY, assign_cathy=True)  # Day B
    staff = _staff_by_client(ws)

    assert staff["Iop, Ida"] == "Cathy"
    assert staff["Iop, Ivan"] == "Cathy"
    # Professional bills every day, so Wednesday Detox/e-care are hers too.
    assert staff["Detox, Dora"] == "Cathy"
    assert staff["Ecare, Ellis"] == "Cathy"

    # The two pool rows split one each; Day B gives the first half
    # ("Iop, Inst" sorts before "Iop, Otto") to Cathy.
    assert staff["Iop, Inst"] == "Cathy"
    assert staff["Iop, Otto"] == "Jasmine"


def _cathy_all_payer_rows():
    """Professional rows for the payers only Cathy's full list covers."""
    return [
        ("Insurance", "Individual Therapy", "Emblem (Optum)", "Gold, Gil", "CMS-1500"),
        ("Insurance", "Individual Therapy", "Surest (Optum)", "Hall, Hana", "UB-04"),
        ("Insurance", "Individual Therapy", "UMR (Optum)", "Ives, Ike", "CMS-1500"),
        ("Insurance", "Individual Therapy", "UBH-HP (Optum)", "Jones, Jo", "CMS-1500"),
    ]


def test_cathy_all_payers_turns_the_carveout_on_by_itself():
    """cathy_all_payers alone runs the carve-out; assign_cathy is not needed."""
    ws = _sheet(_cathy_all_payer_rows())
    assign_staff(ws, "09232026", cathy_all_payers=True)  # Day A
    # Without the carve-out, Day A would give Gold/Hall to Jasmine.
    assert set(_assignments(ws)) == {"Cathy"}


def test_cathy_all_payers_rows_leave_the_shared_pool():
    """Her wider payer list shrinks the pool the split is applied to."""
    rows = _pool_rows(40, prefix="Cathy", payer="Emblem (Optum)")
    rows += _pool_rows(160)

    ws = _sheet(rows)
    assign_staff(ws, WEDNESDAY, cathy_all_payers=True)
    assignments = _assignments(ws)

    # 40 carve-out rows plus half of the 160 left in the pool.
    assert assignments.count("Cathy") == 120
    assert assignments.count("Jasmine") == 80
    # 200 rows in, 200 rows out, each with exactly one owner.
    assert len(assignments) == 200
    assert all(value for value in assignments)


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


def test_skip_cathy_also_disables_her_carveout():
    """skip_cathy turns off the payer carve-out too: she gets nothing at all."""
    ws = _sheet([
        ("Insurance", "IOP", "Oxford", "Iop, Ida", "CMS-1500"),
    ])
    assign_staff(ws, WEDNESDAY, assign_cathy=True, skip_cathy=True)
    assert _staff_by_client(ws)["Iop, Ida"] == "Jasmine"


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
    assign_staff(ws, WEDNESDAY, include_programming=True, assign_cathy=True)
    staff = _staff_by_client(ws)

    assert staff["Aetna, Amy"] == "Melissa"
    assert staff["Humana, Hal"] == "Melissa"


def test_split_applies_to_the_pool_left_after_her_carveout():
    """Cathy's carve-out rows are removed before the pool is split in half."""
    rows = _pool_rows(20, prefix="Carveout", payer="Oxford")
    rows += _pool_rows(20)

    ws = _sheet(rows)
    assign_staff(ws, WEDNESDAY, assign_cathy=True)
    assignments = _assignments(ws)

    # 20 carve-out rows plus half of the remaining 20-row pool.
    assert assignments.count("Cathy") == 30
    assert assignments.count("Jasmine") == 10


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
    """A custom report claims Professional rows for its payers, like a second Cathy."""
    ws = _sheet(_cathy_candidate_rows())
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
    """Custom report rows leave the shared pool entirely, same as Cathy's carve-out."""
    rows = _pool_rows(40, prefix="Karen", payer="Cigna")
    rows += _pool_rows(160, payer="Optum")

    ws = _sheet(rows)
    assign_staff(ws, WEDNESDAY, custom_report_name="Karen",
                 custom_report_payer_terms=["cigna"])
    assignments = _assignments(ws)

    assert assignments.count("Karen") == 40
    assert assignments.count("Cathy") == 80
    assert assignments.count("Jasmine") == 80


def test_custom_report_and_cathy_do_not_double_claim():
    """When both are on, the carve-out is checked first; the custom report never re-claims its rows."""
    rows = [
        ("Insurance", "Individual Therapy", "Oxford", "Adams, Ann", "CMS-1500"),
        ("Insurance", "Individual Therapy", "Cigna", "Baker, Bob", "CMS-1500"),
    ]
    ws = _sheet(rows)
    # A custom report configured to also match Oxford: the carve-out still
    # gets it, because assign_cathy is checked first in assign_staff.
    assign_staff(ws, WEDNESDAY, assign_cathy=True,
                 custom_report_name="Karen",
                 custom_report_payer_terms=["oxford", "cigna"])
    staff = _staff_by_client(ws)

    assert staff["Adams, Ann"] == "Cathy"
    assert staff["Baker, Bob"] == "Karen"


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
