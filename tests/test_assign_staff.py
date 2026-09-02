"""
End-to-end tests for assign_staff: the Cathy report and the
include_programming override, exercised against a real worksheet.

assign_staff lives in both app.py and "Unbilled Step 1.py". app.py imports
Streamlit, so these tests load the standalone script instead — the two
copies of the function are kept in sync.

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

# MMDDYYYY tokens for the two days that matter to these tests.
WEDNESDAY = "09022026"
TUESDAY = "09012026"


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
        # Not Cathy's: right payer, but not a Professional claim type.
        ("Insurance", "Individual Therapy", "Oxford", "Diaz, Dee", "837I"),
        # Not Cathy's: Professional, but a payer she does not cover.
        ("Insurance", "Individual Therapy", "Aetna", "Evans, Eve", "CMS-1500"),
        ("Insurance", "Individual Therapy", "Optum", "Frank, Fay", "CMS-1500"),
    ]


def test_cathy_off_by_default():
    """Without the flag, nobody is assigned to Cathy and the pool is unchanged."""
    ws = _sheet(_cathy_candidate_rows())
    assign_staff(ws, WEDNESDAY)
    staff = _staff_by_client(ws)

    assert "Cathy" not in staff.values()
    # Every Professional row falls to Rosanna under her weekday cap.
    for client in ("Adams, Ann", "Baker, Bob", "Carter, Cal", "Evans, Eve",
                   "Frank, Fay"):
        assert staff[client] == "Rosanna", (client, staff[client])


def test_cathy_takes_only_professional_rows_for_her_payers():
    """Oxford/ConnectiCare/UBH Professional rows go to Cathy; nothing else does."""
    ws = _sheet(_cathy_candidate_rows())
    assign_staff(ws, WEDNESDAY, assign_cathy=True)
    staff = _staff_by_client(ws)

    assert staff["Adams, Ann"] == "Cathy"
    assert staff["Baker, Bob"] == "Cathy"
    assert staff["Carter, Cal"] == "Cathy"

    # 837I Oxford is not Professional, so it is not Cathy's; on a Wednesday
    # it is not billable either.
    assert staff["Diaz, Dee"] == "Unable to Bill"
    # Professional rows for other payers stay in the Rosanna/Jasmine pool.
    assert staff["Evans, Eve"] == "Rosanna"
    assert staff["Frank, Fay"] == "Rosanna"


def test_cathy_rows_leave_the_professional_pool():
    """A row assigned to Cathy is not also given to Rosanna or Jasmine."""
    ws = _sheet(_cathy_candidate_rows())
    assign_staff(ws, WEDNESDAY, assign_cathy=True)

    assignments = [ws.cell(row, 1).value for row in range(2, ws.max_row + 1)]
    assert assignments.count("Cathy") == 3
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


def test_cathy_takes_iop_for_her_payers():
    """Every Professional service for Cathy's payers is hers, IOP included."""
    ws = _sheet([
        ("Insurance", "IOP", "Oxford", "Iop, Ida", "CMS-1500"),
        ("Insurance", "Telemed IOP", "ConnectiCare", "Iop, Ivan", "UB-04"),
        ("Insurance", "Detox Admission", "UBH", "Detox, Dora", "CMS-1500"),
        ("Insurance", "E-Care Individual", "Oxford", "Ecare, Ellis", "CMS-1500"),
        # IOP for a payer Cathy does not cover is still Jasmine's.
        ("Insurance", "Telemed IOP", "Optum", "Iop, Otto", "CMS-1500"),
        # IOP for one of her payers, but not a Professional claim type, is
        # still Jasmine's.
        ("Insurance", "IOP", "Oxford", "Iop, Inst", "837I"),
    ])
    assign_staff(ws, WEDNESDAY, assign_cathy=True)
    staff = _staff_by_client(ws)

    assert staff["Iop, Ida"] == "Cathy"
    assert staff["Iop, Ivan"] == "Cathy"
    # Professional bills every day, so Wednesday Detox/e-care are hers too.
    assert staff["Detox, Dora"] == "Cathy"
    assert staff["Ecare, Ellis"] == "Cathy"

    assert staff["Iop, Otto"] == "Jasmine"
    assert staff["Iop, Inst"] == "Jasmine"

    # Her rows are labelled "Cathy" in the Staff/Status column, so the
    # Masters workbook names her as the owner too.
    assignments = [ws.cell(row, 1).value for row in range(2, ws.max_row + 1)]
    assert assignments.count("Cathy") == 4


def test_programming_included_on_a_wednesday():
    """include_programming makes Wednesday Detox/Residential billable to Jasmine."""
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
    assign_staff(ws, WEDNESDAY, include_programming=True)
    staff = _staff_by_client(ws)
    assert staff["Detox, Dan"] == "Jasmine"
    assert staff["Res, Rita"] == "Jasmine"
    # e-care is untouched by the override: still Tuesday-only.
    assert staff["Ecare, Ed"] == "Unable to Bill"

    ws = _sheet(rows)
    assign_staff(ws, TUESDAY, include_programming=True)
    staff = _staff_by_client(ws)
    assert staff["Ecare, Ed"] == "Jasmine"


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


def test_rosanna_cap_applies_to_the_pool_left_after_cathy():
    """Cathy's rows are removed before Rosanna's 150-row cap is applied."""
    rows = [("Insurance", "Individual Therapy", "Oxford", f"Cathy{i:04d}", "CMS-1500")
            for i in range(40)]
    rows += [("Insurance", "Individual Therapy", "Optum", f"Pool{i:04d}", "CMS-1500")
             for i in range(160)]

    ws = _sheet(rows)
    assign_staff(ws, WEDNESDAY, assign_cathy=True)
    assignments = [ws.cell(row, 1).value for row in range(2, ws.max_row + 1)]

    assert assignments.count("Cathy") == 40
    # 160 rows are left in the pool: Rosanna's 150, then Jasmine's 10.
    assert assignments.count("Rosanna") == 150
    assert assignments.count("Jasmine") == 10


def _dropdown_options(**kwargs):
    """Run finalize_workbook on a small workbook and read back its Status list."""
    wb = Workbook()
    ws = wb.active
    ws.append(HEADERS)
    ws.append(["Rosanna", "OP Westchester", "Insurance", "Individual Therapy",
               "Oxford", "Smith, John", "OP", "Adams, Ann", "CMS-1500"])
    unbilled.finalize_workbook(wb, **kwargs)
    sheet2 = wb["Sheet2"]
    return [sheet2.cell(row, 1).value for row in range(1, sheet2.max_row + 1)]


def test_cathy_status_dropdown_matches_jasmine():
    """Cathy's Status dropdown carries the same options as Jasmine's."""
    base = ["Billed", "Unable to Bill", "Contractual Adj", "Incomplete Billings",
            "Utox Batch", "Inclusive Services"]

    # Rosanna's list: the six shared options.
    assert _dropdown_options() == base

    # Jasmine's and Cathy's list: the same six plus Batch Billings and IOP.
    jasmine = _dropdown_options(include_batch_billings=True, include_iop_status=True)
    assert jasmine == base + ["Batch Billings", "IOP"]


if __name__ == '__main__':
    for name, test in sorted(
        (name, obj) for name, obj in list(globals().items())
        if name.startswith('test_') and callable(obj)
    ):
        test()
        print(f"✓ {name} passed")

    print("\nAll tests passed!")
