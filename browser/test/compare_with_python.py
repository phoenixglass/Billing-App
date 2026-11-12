"""Check the browser version against the Python app, cell by cell.

For every test case, the same input workbook goes through app.process_workbook
(the real Streamlit app's processing) and through the browser version's code
in Node (test/run.js). Every output workbook is compared: file names, sheet
names, every cell value, the Status dropdown's cells, and column widths in the
individual reports. Inputs are synthetic (invented clients) built from the
real export's columns and category values; no patient data is involved.

Run from the repo root, with the app's requirements installed and
`npm install` done in browser/:
    python browser/test/compare_with_python.py [--demo DIR_WITH_DEMO_FILES]
"""
import argparse
import datetime as dt
import io
import json
import os
import random
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl

REPO = Path(__file__).resolve().parents[2]
BROWSER = REPO / "browser"

HEADERS = [
    "GROUPFLD1", "GROUPFLD2", "Client", "Payer", "Service", "Range", "Amount", "Services",
    "Claim Type", "MRN", "Units", "Duration", "Service Category", "Facility As Rendering",
    "Billing Provider", "Rendering Provider", "Attending Provider", "Pay Order", "Program Level",
    "Account", "Missing Required Note", "Requires Review", "Ready To Bill", "Billable Status",
    "CPT", "Rev Code", "Modifier", "Bill Type", "POS", "Authorization", "Auth Start", "Auth End",
    "Episode Start", "Episode End", "Episode Type", "Insured DOB", "Policy #", "Insured DOB",
    "Insured Gender", "Diag 1", "Diag 2", "Active Program", "Group Service", "Billing Batch",
    "Batch Last Run", "Min Created", "Max Created", "EMG", "COB", "ICN# for Corrected Claim",
    "Service Note", "Facility Name as Billing Provider", "CMS BOX 24H",
    "Use Facility NPI as Rendering", "Progress Note ID", "NDC Value", "NDC Measurement", "NDC QTY",
]
SERVICES = """Acupuncture - Add'l 15 mins|Acupuncture- Self-Pay OPS|Case Management|Crisis Psychotherapy 30-60min|Detox|Detox Scholarship|Drug Screen 13 Panel OPS|Drug Screen 13 Panel OPS- Group|E-CARE Housing Cost|E-care PHP Housing|E-Care Scholarship|EKG Performance|IOP Canaan|IOP NYC|IOP-Wilton|Medication Admin/Injection|No Show: PHP Group|Nutrition Group|Outpatient 53+|Outpatient Group (75-90 minutes)|Outpatient Withdrawal Management (OWM Appeal)|Partial Hospitalization|PHP Census Tracking|PHP: Virtual|Residential 3.5|Residential Suite|Residential: APPEAL|SBI 30 +|Telemed: Day Treatment Mass|Telemed: IOP|Telemed: Outpatient 53+ minutes|Telemed: Outpatient Group - Adolescent""".split("|")
PAYERS = ["Cigna", "Aetna", "Aetna -", "Humana", "Optum", "Surest (Optum)", "BCBS - Mass CT", "BCBS",
          "BCB Anthem CT", "Anthem BCBS", "Beacon", "Oxford", "UnitedHealthcare", "Self Pay"]
DIVS = ["CA - PHP", "Detox", "Extended Care", "General Charges", "OP Canaan", "OP Chappaqua",
        "OP Huntington", "OP Massachussets", "OP NYC", "OP Ramsey", "OP Wilton", "Residential"]
LEVELS = ["Ca-IOP", "Ca-PHP", "Ch-OP", "Ch-Virtual OP WM", "D-Detox", "E-Men", "H-IOP",
          "M-Virtual Day Treatment", "NYC-OP", "NYC-OP WM", "NYC-PHP", "R-Residential", "W-OP", None]
CLAIMS = ["837I", "CMS-1500", "CMS-1500", "Standard Invoice", "UB-04", None]
GROUP2 = ["Insurance"] * 7 + ["Self Pay", None]
PROVS = ["Demo Clinician, Alex", "Demo Clinician, Sam", "O'Flynn, Karen"]
STATUS = ["Valid for Billing"] * 8 + ["Invalid for Billing", None]
FIRST = ["Ann", "Bo", "Cy", "Di", "Ed"]
SURNAMES = ["O'Brien", "Oakes", "Smith-Jones", "Smith", "smithers", "McAdams", "Mcadams", "de la Cruz",
            "DeLacy", "Zane", "Abbott", "abbot"]

# Option sets: (browser options, matching app.process_workbook keyword arguments)
OPTION_SETS = [
    ({}, {}),
    ({"excludeOptum": True, "excludeBcbAnthemCt": True, "excludeAnthemCathyJasmine": True},
     {"exclude_optum": True, "exclude_bcb_anthem_ct": True, "exclude_anthem_cathy_jasmine_owm": True}),
    ({"excludeAetna": True, "excludeDetoxResidential": True, "includeProgramming": True},
     {"exclude_aetna": True, "exclude_detox_residential": True, "include_programming": True}),
    ({"includeProgramming": True}, {"include_programming": True}),
    ({"skipCathy": True, "phpToJasmine": True}, {"skip_cathy": True, "php_to_jasmine": True}),
    ({"splitDay": "B"}, {"split_day": "B"}),
    ({"splitDay": "A", "phpToJasmine": True}, {"split_day": "A", "php_to_jasmine": True}),
    ({"customReportName": "Karen", "customReportPayerTerms": ["humana", "cigna"]},
     {"custom_report_name": "Karen", "custom_report_payer_terms": ["humana", "cigna"]}),
    ({"customReportName": "Rosanna", "customReportPayerTerms": ["optum"], "customReportProfessionalOnly": False},
     {"custom_report_name": "Rosanna", "custom_report_payer_terms": ["optum"], "custom_report_professional_only": False}),
    ({"payerTerms": ["oxford"], "serviceTerms": ["group"], "scope": ["Jasmine"]},
     {"exclude_payer_terms": ["oxford"], "exclude_service_terms": ["group"], "exclude_scope": ["Jasmine"]}),
    ({"payerTerms": ["cigna", "beacon"], "scope": ["Cathy", "CB"]},
     {"exclude_payer_terms": ["cigna", "beacon"], "exclude_scope": ["Cathy", "CB"]}),
    ({"serviceTerms": ["outpatient"]}, {"exclude_service_terms": ["outpatient"]}),
    ({"divisionPayerTerms": ["BCBS", "Beacon"], "divisionTerms": ["Residential", "Detox"]},
     {"division_exclude_payer_terms": ["BCBS", "Beacon"], "division_exclude_division_terms": ["Residential", "Detox"]}),
    ({"excludeAetna": True, "skipCathy": True, "payerTerms": ["humana"], "scope": ["Cathy"],
      "customReportName": "Karen", "customReportPayerTerms": ["oxford"], "divisionPayerTerms": ["cigna"],
      "divisionTerms": ["wilton"], "splitDay": "B", "includeProgramming": True},
     {"exclude_aetna": True, "skip_cathy": True, "exclude_payer_terms": ["humana"], "exclude_scope": ["Cathy"],
      "custom_report_name": "Karen", "custom_report_payer_terms": ["oxford"],
      "division_exclude_payer_terms": ["cigna"], "division_exclude_division_terms": ["wilton"],
      "split_day": "B", "include_programming": True}),
]


def random_report(path, seed, n, report_date):
    rnd = random.Random(seed)
    col = {h: i for i, h in reversed(list(enumerate(HEADERS)))}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(HEADERS)
    for i in range(n):
        r = [None] * len(HEADERS)
        r[col["Client"]] = (f"{rnd.choice(SURNAMES)}, {rnd.choice(FIRST)}" if rnd.random() < 0.4 else
                            rnd.choice("ABCDEFGHIJKLMNOPQRSTUVWXYZ") + "".join(rnd.choice("aeioulmnrst") for _ in range(rnd.randint(2, 7)))
                            + ", " + rnd.choice(FIRST))
        r[col["Service"]] = rnd.choice(SERVICES)
        r[col["Payer"]] = rnd.choice(PAYERS)
        r[col["GROUPFLD1"]] = rnd.choice(DIVS)
        r[col["GROUPFLD2"]] = rnd.choice(GROUP2)
        r[col["Program Level"]] = rnd.choice(LEVELS)
        r[col["Claim Type"]] = rnd.choice(CLAIMS)
        r[col["Billing Provider"]] = rnd.choice(PROVS)
        r[col["Billable Status"]] = rnd.choice(STATUS)
        r[col["Range"]] = dt.datetime.combine(report_date - dt.timedelta(days=1 + i % 3), dt.time())
        r[col["Amount"]] = rnd.choice([75, 120, 185, 410.5, 1450])
        r[col["MRN"]] = f"DEMO-MRN-{seed:02d}{i:04d}"
        r[col["Units"]] = 1
        ws.append(r)
    for (c,) in ws.iter_rows(min_row=2, min_col=col["Range"] + 1, max_col=col["Range"] + 1):
        c.number_format = "mm/dd/yyyy"
    ws.append([None] * len(HEADERS))
    ws.cell(ws.max_row, col["GROUPFLD1"] + 1, "Total")
    wb.save(path)


def load_python_app():
    home = Path(tempfile.mkdtemp())
    (home / ".streamlit").mkdir()
    (home / ".streamlit" / "secrets.toml").write_text("")
    os.environ["HOME"] = str(home)
    os.chdir(home)  # audit_logs/ is created in the working directory
    sys.path.insert(0, str(REPO))
    import contextlib
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        import app  # noqa: E402  (Streamlit runs in "bare" mode: the UI calls are no-ops)
    import logging
    logging.disable(logging.CRITICAL)
    return app


class Upload:
    def __init__(self, path, name):
        self.name = name
        self._data = Path(path).read_bytes()

    def getbuffer(self):
        return memoryview(self._data)

    def getvalue(self):
        return self._data


def norm(v):
    if isinstance(v, dt.datetime):
        return v.replace(tzinfo=None).isoformat(timespec="seconds")
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if v == "":
        return None
    return v


def grid(ws):
    rows = [[norm(c) for c in r] for r in ws.iter_rows(values_only=True)]
    while rows and all(v is None for v in rows[-1]):
        rows.pop()
    width = max((max((i + 1 for i, v in enumerate(r) if v is not None), default=0) for r in rows), default=0)
    return [r[:width] + [None] * (width - len(r)) for r in rows]


def widths(ws):
    """Column number -> width. A <col> entry can cover several columns
    (ExcelJS groups equal neighbours), so expand each entry's min..max."""
    out = {}
    for dim in ws.column_dimensions.values():
        if dim.width is not None and dim.min:
            for c in range(dim.min, (dim.max or dim.min) + 1):
                out[c] = round(dim.width, 1)
    return out


def dv_cells(ws):
    cells = set()
    for d in ws.data_validations.dataValidation:
        for rng in str(d.sqref).split():
            for row in ws[rng] if ":" in rng else [[ws[rng]]]:
                for c in row:
                    cells.add(c.coordinate)
    return cells


def compare_files(py_bytes, js_path, name, is_report):
    problems = []
    a = openpyxl.load_workbook(io.BytesIO(py_bytes))
    b = openpyxl.load_workbook(js_path)
    if a.sheetnames != b.sheetnames:
        return [f"{name}: sheets {a.sheetnames} vs {b.sheetnames}"]
    for sn in a.sheetnames:
        ga, gb = grid(a[sn]), grid(b[sn])
        if ga != gb:
            diffs = [(i + 1, x, y) for i, (x, y) in enumerate(zip(ga, gb)) if x != y][:2]
            problems.append(f"{name}/{sn}: {len(ga)} vs {len(gb)} rows; first diffs {diffs}")
        if is_report and sn == "Sheet1":
            if dv_cells(a[sn]) != dv_cells(b[sn]):
                problems.append(f"{name}: dropdown cells differ")
            wa, wb_ = widths(a[sn]), widths(b[sn])
            cols = range(1, a[sn].max_column + 1)
            if [wa.get(i) for i in cols] != [wb_.get(i) for i in cols]:
                bad = [(openpyxl.utils.get_column_letter(i), wa.get(i), wb_.get(i)) for i in cols
                       if wa.get(i) != wb_.get(i)][:3]
                problems.append(f"{name}: column widths differ, e.g. {bad}")
    return problems


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--demo", help="folder with the DEMO Unbilled Revenue *.xlsx files")
    ap.add_argument("--seeds", type=int, default=8)
    args = ap.parse_args()

    work = Path(tempfile.mkdtemp())
    inputs = []
    if args.demo:
        inputs += [(p, p.name) for p in sorted(Path(args.demo).glob("DEMO Unbilled Revenue *.xlsx"))]
    start = dt.date(2026, 11, 16)  # a Monday
    for seed in range(args.seeds):
        d = start + dt.timedelta(days=seed)
        name = f"Unbilled Revenue By Resident and Funding Type {d:%m%d%Y}.xlsx"
        p = work / f"in_{seed}" / name
        p.parent.mkdir()
        random_report(p, seed, 160, d)
        inputs.append((p, name))

    app = load_python_app()
    cases = failures = rows_checked = 0
    for path, name in inputs:
        for js_opts, py_opts in OPTION_SETS:
            cases += 1
            outdir = work / f"js_{cases}"
            subprocess.run(["node", str(BROWSER / "test" / "run.js"), str(path), str(outdir),
                            json.dumps(js_opts), name], check=True, capture_output=True)
            js = json.loads((outdir / "summary.json").read_text())
            files, invalid, token, wm, split = app.process_workbook(Upload(path, name), **py_opts)
            problems = []
            if not js["ok"]:
                problems.append(f"browser version failed: {js['error']}")
            else:
                if sorted(files) != sorted(js["files"]):
                    problems.append(f"files {sorted(files)} vs {sorted(js['files'])}")
                summary_py = (invalid, token, wm, split)
                summary_js = (js["invalidCount"], js["dateToken"], js["wmCount"], js["splitDay"])
                if summary_py != summary_js:
                    problems.append(f"summary {summary_py} vs {summary_js}")
                for fname, buf in files.items():
                    if (outdir / fname).exists():
                        problems += compare_files(buf.getvalue(), outdir / fname, fname,
                                                  is_report=not fname.endswith("Masters.xlsx"))
                        rows_checked += openpyxl.load_workbook(io.BytesIO(buf.getvalue())).worksheets[0].max_row - 1
            status = "ok" if not problems else "MISMATCH"
            failures += bool(problems)
            print(f"{status:8} {name[-13:-5]} {json.dumps(js_opts)[:90]}")
            for p in problems[:4]:
                print("         ", p)
    print(f"\n{cases} cases, {rows_checked} output rows compared, {failures} with differences")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
