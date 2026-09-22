import openpyxl
from pathlib import Path
import re
from datetime import datetime

from billing_rules import (
    is_aetna_payer,
    is_wm_program_level,
    CATHY_PAYERS,
    CATHY_ALL_PAYERS,
    validate_custom_report_name,
    parse_terms,
    matches_any_term,
    payer_excluded_by_division,
    assign_staff,
    finalize_workbook,
)

stubs_required = False  # no Streamlit in the standalone script; keep for parity if needed

def extract_date_from_filename(filename):
    """Extract MMDDYYYY date from filename"""
    match = re.search(r'(\d{8})', filename)
    if match:
        return match.group(1)
    return None

def get_save_folder(workbook_path, date_token):
    """Build save path: ...\\Unbilled Reports\\YYYY\\MM - Month YYYY\\MMDDYYYY"""
    wb_path = Path(workbook_path)

    # Find "Unbilled Reports" in the path
    parts = wb_path.parts
    try:
        idx = [i for i, p in enumerate(parts) if 'Unbilled Reports' in p][0]
        base = Path(*parts[:idx+1])
    except IndexError:
        raise ValueError("Workbook must be saved in a folder containing 'Unbilled Reports'")

    # Parse date token: MMDDYYYY
    month = date_token[:2]
    day = date_token[2:4]
    year = date_token[4:]

    dt = datetime.strptime(date_token, '%m%d%Y')
    month_name = dt.strftime('%B')

    # Build: YYYY\\MM - Month YYYY\\MMDDYYYY
    save_path = base / year / f"{month} - {month_name} {year}" / date_token
    save_path.mkdir(parents=True, exist_ok=True)

    return save_path

def step_1_extract_invalid(ws):
    """Insert Staff/Status, extract Invalid rows to new sheet, delete from main"""

    # Insert Staff/Status column if needed
    if ws.cell(1, 1).value != "Staff/Status":
        ws.insert_cols(1)
        ws.cell(1, 1).value = "Staff/Status"

    # Find Billable Status column
    billable_col = None
    for col in range(1, ws.max_column + 1):
        if ws.cell(1, col).value == "Billable Status":
            billable_col = col
            break

    if not billable_col:
        raise ValueError("Billable Status column not found")

    # Find all invalid rows
    invalid_rows = []
    for row in range(2, ws.max_row + 1):
        if ws.cell(row, billable_col).value == "Invalid for Billing":
            invalid_rows.append(row)

    if not invalid_rows:
        print("No invalid rows found")
        return None

    # Create Invalid sheet
    wb = ws.parent
    if "Invalid" in wb.sheetnames:
        del wb["Invalid"]

    ws_invalid = wb.create_sheet("Invalid")

    # Copy header
    for col in range(1, ws.max_column + 1):
        ws_invalid.cell(1, col).value = ws.cell(1, col).value

    # Copy invalid rows
    for idx, row_num in enumerate(invalid_rows, start=2):
        for col in range(1, ws.max_column + 1):
            ws_invalid.cell(idx, col).value = ws.cell(row_num, col).value

    # Delete invalid rows (backwards to preserve row numbers)
    for row_num in reversed(invalid_rows):
        ws.delete_rows(row_num)

    # Format headers
    ws.cell(1, 1).font = openpyxl.styles.Font(bold=True)
    ws.auto_filter.ref = ws.dimensions

    print(f"Extracted {len(invalid_rows)} invalid rows to 'Invalid' sheet")
    return len(invalid_rows)

def export_staff_workbooks(wb, wb_path, date_token, exclude_aetna: bool = False,
                           skip_cathy: bool = False,
                           exclude_payer_terms: list = None,
                           exclude_service_terms: list = None,
                           exclude_scope: list = None,
                           division_exclude_payer_terms: list = None,
                           division_exclude_division_terms: list = None,
                           custom_report_name: str = None):
    """Export separate workbooks for Cathy, Jasmine, and CB.

    All staff (Melissa, Unable to Bill, etc.) are still assigned in the
    Masters workbook, but only Cathy, Jasmine, and CB receive individual
    reports. Cathy's workbook is always exported (like Jasmine/CB) unless
    skip_cathy is set, since her carve-out and pool share both land in her
    own "Cathy" assignment.

    Args:
        exclude_aetna: When True, Aetna rows are left out of every individual
            workbook. They still appear in the Masters workbook.
        skip_cathy: When True, no workbook is exported for Cathy — she was
            assigned nothing at all for this run.
        exclude_payer_terms/exclude_service_terms: free-text custom
            exclusions (--exclude-payers/--exclude-services). Any row whose
            Payer or Service contains one of these terms (case-insensitive)
            is left out of the individual workbooks for this run only.
        exclude_scope: staff names (--exclude-scope) the two custom
            exclusions apply to. Empty/None applies them to everyone, same
            as exclude_aetna.
        division_exclude_payer_terms/division_exclude_division_terms:
            free-text funding-source-by-division exclusion
            (--exclude-payer-by-division-payers/--exclude-payer-by-division-divisions).
            A row is left out of every individual workbook only when its
            Payer matches one of division_exclude_payer_terms AND its
            GROUPFLD1 matches one of division_exclude_division_terms.
        custom_report_name: when set, also export this staff member's
            workbook (the rows assign_staff routed to them via
            --custom-report-name/--custom-report-payers).
    """

    ws = wb.active
    save_folder = get_save_folder(wb_path, date_token)

    # Find Program Level column for WM filtering, Payer for the Aetna/custom
    # exclusions, Service for the custom service exclusion, and GROUPFLD1
    # (division) for the funding-source-by-division exclusion.
    program_level_col = None
    payer_col = None
    service_col = None
    division_col = None
    for col in range(1, ws.max_column + 1):
        header = ws.cell(1, col).value
        if header == "Program Level":
            program_level_col = col
        elif header == "Payer":
            payer_col = col
        elif header == "Service":
            service_col = col
        elif header == "GROUPFLD1":
            division_col = col

    # Count and warn about WM rows
    wm_count = 0
    if program_level_col is not None:
        for row in range(2, ws.max_row + 1):
            if is_wm_program_level(ws.cell(row, program_level_col).value):
                wm_count += 1
    if wm_count > 0:
        print(
            f"WARNING: {wm_count} row(s) with 'WM' in the Program Level column were found. "
            "These rows are assigned to Melissa and appear only in the Masters report. "
            "Only Melissa is authorized to bill WM."
        )

    staff_reports = ["Jasmine", "CB"] if skip_cathy else ["Cathy", "Jasmine", "CB"]
    if custom_report_name:
        staff_reports.append(custom_report_name)

    for staff_name in staff_reports:
        # Create new workbook
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active
        new_ws.title = "Sheet1"

        # Copy header
        for col in range(1, ws.max_column + 1):
            new_ws.cell(1, col).value = ws.cell(1, col).value

        # Copy matching rows
        new_row = 2
        for row in range(2, ws.max_row + 1):
            assigned_staff = ws.cell(row, 1).value
            if assigned_staff == staff_name:
                # Exclude Aetna rows from every individual workbook; they are
                # still assigned in the Masters workbook.
                if exclude_aetna and payer_col is not None:
                    if is_aetna_payer(str(ws.cell(row, payer_col).value or "")):
                        continue
                # Remove a funding source (Payer) by division (GROUPFLD1):
                # only excluded when both a payer term and a division term
                # are given and both match this row.
                if (division_exclude_payer_terms and division_exclude_division_terms
                        and payer_col is not None and division_col is not None):
                    payer_val = str(ws.cell(row, payer_col).value or "")
                    division_val = str(ws.cell(row, division_col).value or "")
                    if payer_excluded_by_division(
                            payer_val, division_val,
                            division_exclude_payer_terms, division_exclude_division_terms):
                        continue
                # Custom per-run exclusions (free text, no code change
                # needed): skip if this staff is in scope (or scope is
                # empty, meaning everyone) and the payer/service matches.
                in_scope = not exclude_scope or staff_name in exclude_scope
                if in_scope and exclude_payer_terms and payer_col is not None:
                    if matches_any_term(str(ws.cell(row, payer_col).value or ""), exclude_payer_terms):
                        continue
                if in_scope and exclude_service_terms and service_col is not None:
                    if matches_any_term(str(ws.cell(row, service_col).value or ""), exclude_service_terms):
                        continue
                # Skip WM program level rows for all staff except Melissa
                # (Masters retains all rows; only Melissa bills WM)
                if (staff_name != "Melissa" and
                        program_level_col is not None and
                        is_wm_program_level(ws.cell(row, program_level_col).value)):
                    continue
                for col in range(1, ws.max_column + 1):
                    new_ws.cell(new_row, col).value = ws.cell(row, col).value
                new_row += 1

        if new_row == 2:  # No rows found
            print(f"No rows for {staff_name}, skipping")
            continue

        # Finalize for Cathy, Jasmine, and a custom report's staff, all with
        # the wider dropdown (Batch Billings and IOP included). CB is never
        # in staff_reports' finalized set (it has no status dropdown).
        if staff_name != "CB":
            finalize_workbook(new_wb, include_batch_billings=True,
                              include_iop_status=True)

        # Save
        save_path = save_folder / f"Unbilled Revenue By Resident and Funding Type {date_token} - {staff_name}.xlsx"
        new_wb.save(save_path)
        print(f"Saved {save_path}")

def main(workbook_path, include_programming: bool = False, exclude_aetna: bool = False,
         cathy_report: bool = False, cathy_all_payers: bool = False,
         skip_cathy: bool = False, even_split_jasmine_cathy: bool = False,
         exclude_payer_terms: list = None,
         exclude_service_terms: list = None, exclude_scope: list = None,
         cathy_cap_override: int = None, custom_report_name: str = None,
         custom_report_payer_terms: list = None,
         custom_report_professional_only: bool = True,
         division_exclude_payer_terms: list = None,
         division_exclude_division_terms: list = None):
    """Main workflow.

    The optional flags mirror the checkboxes in the Streamlit app and
    are all off by default, so a plain run follows the standard daily
    schedule:
      include_programming - bill Programming (Detox/Residential) regardless
          of the weekday.
      exclude_aetna       - keep Aetna rows out of the individual workbooks.
      cathy_report        - route Professional Oxford/ConnectiCare/UBH rows
          to Cathy on top of her standing share of the professional pool.
      cathy_all_payers    - run that carve-out against her full payer list
          (CATHY_ALL_PAYERS) instead of her usual three; turns the carve-out
          on by itself.
      skip_cathy          - give Cathy nothing at all: Jasmine takes the
          whole professional pool and no workbook is saved for Cathy.
      even_split_jasmine_cathy - instead of Cathy's cap and Jasmine taking
          the remainder, split everything Jasmine would otherwise get
          50/50 between Jasmine and Cathy for this run (IOP still always
          goes to Jasmine). Ignored if skip_cathy is set.
      exclude_payer_terms/exclude_service_terms - free-text custom
          exclusions (--exclude-payers/--exclude-services): rows whose
          Payer/Service contains any of these terms are left out of the
          individual workbooks for this run only.
      exclude_scope       - staff names (--exclude-scope) the two custom
          exclusions apply to; empty/None applies them to everyone.
      cathy_cap_override  - give Cathy exactly this many professional-pool
          rows for this run instead of the standard weekday schedule
          (--cathy-cap). Ignored if skip_cathy or even_split_jasmine_cathy
          is set.
      custom_report_name/custom_report_payer_terms/
          custom_report_professional_only - a second, generic Cathy-shaped
          report (--custom-report-name/--custom-report-payers/
          --custom-report-any-claim-type): matching rows go to a new named
          staff member with their own workbook.
      division_exclude_payer_terms/division_exclude_division_terms - remove
          a funding source (Payer) by division (--exclude-payer-by-division-payers/
          --exclude-payer-by-division-divisions), each a comma-separated
          list so several funding sources and several divisions can be
          combined in one run. A row is left out of every individual
          workbook only when its Payer matches one of the payer terms AND
          its GROUPFLD1 (division) matches one of the division terms.
    """
    # "All of her payers" runs the Cathy report on its own.
    cathy_report = cathy_report or cathy_all_payers
    if custom_report_name:
        validate_custom_report_name(custom_report_name)

    # Load workbook
    wb = openpyxl.load_workbook(workbook_path)

    # Find the data sheet (first sheet or by name pattern)
    ws = wb.active

    # Extract date token
    date_token = extract_date_from_filename(Path(workbook_path).name)
    if not date_token:
        print("Warning: Could not extract date (MMDDYYYY) from filename, will use today's date")

    print(f"Processing file with date: {date_token if date_token else 'today'}")

    # Step 1: Extract invalid
    step_1_extract_invalid(ws)

    # Step 2-6: Assign staff
    assign_staff(ws, date_token, include_programming=include_programming,
                 assign_cathy=cathy_report, cathy_all_payers=cathy_all_payers,
                 skip_cathy=skip_cathy,
                 cathy_cap_override=cathy_cap_override,
                 even_split_jasmine_cathy=even_split_jasmine_cathy,
                 custom_report_name=custom_report_name,
                 custom_report_payer_terms=custom_report_payer_terms,
                 custom_report_professional_only=custom_report_professional_only)

    # Step 7: Export individual workbooks (only if date_token is available)
    if date_token:
        export_staff_workbooks(wb, workbook_path, date_token,
                               exclude_aetna=exclude_aetna,
                               skip_cathy=skip_cathy,
                               exclude_payer_terms=exclude_payer_terms,
                               exclude_service_terms=exclude_service_terms,
                               exclude_scope=exclude_scope,
                               division_exclude_payer_terms=division_exclude_payer_terms,
                               division_exclude_division_terms=division_exclude_division_terms,
                               custom_report_name=custom_report_name)
    else:
        print("Skipping individual workbook export due to missing date token")

    # Save main workbook
    wb.save(workbook_path)
    print(f"Saved main workbook: {workbook_path}")

    print("\n✓ All done!")

if __name__ == "__main__":
    # Optionally allow running from CLI
    import argparse

    parser = argparse.ArgumentParser(
        description="Process an unbilled revenue workbook."
    )
    parser.add_argument("workbook_path", nargs="?",
                        help="Path to the .xlsx workbook to process.")
    parser.add_argument("--include-programming", action="store_true",
                        help="Bill Programming (Detox/Residential) regardless of "
                             "the weekday, so it can be worked on a Monday or "
                             "Wednesday.")
    parser.add_argument("--exclude-aetna", action="store_true",
                        help="Keep Aetna rows out of the individual staff "
                             "workbooks (they stay in the Masters workbook).")
    parser.add_argument("--cathy-report", action="store_true",
                        help="Assign Professional (CMS-1500/UB-04) Insurance rows "
                             "for " + ", ".join(CATHY_PAYERS) + " to Cathy and "
                             "save her workbook.")
    parser.add_argument("--cathy-all-payers", action="store_true",
                        help="Run the Cathy report against her full payer list ("
                             + ", ".join(CATHY_ALL_PAYERS) + ") instead of just "
                             "her usual three. Turns the Cathy report on by "
                             "itself; --cathy-report is not also needed.")
    parser.add_argument("--no-cathy", action="store_true", dest="skip_cathy",
                        help="Give Cathy nothing at all for this run: Jasmine takes "
                             "the whole Professional pool and no workbook is saved "
                             "for Cathy.")
    parser.add_argument("--even-split-jasmine-cathy", action="store_true",
                        help="Instead of Cathy's cap and Jasmine taking the "
                             "remainder, split everything Jasmine would otherwise "
                             "get 50/50 between Jasmine and Cathy for this run "
                             "(IOP still always goes to Jasmine). Ignored if "
                             "--no-cathy is also set.")
    parser.add_argument("--exclude-payers", default="",
                        help="Comma-separated payer terms (case-insensitive substring "
                             "match). Rows whose Payer contains any of these are left "
                             "out of the individual workbooks for this run only. "
                             "Example: --exclude-payers \"Cigna, Humana\"")
    parser.add_argument("--exclude-services", default="",
                        help="Comma-separated service terms (case-insensitive substring "
                             "match). Rows whose Service contains any of these are left "
                             "out of the individual workbooks for this run only. "
                             "Example: --exclude-services \"Group Therapy\"")
    parser.add_argument("--exclude-scope", default="",
                        help="Comma-separated staff names limiting --exclude-payers/"
                             "--exclude-services to those staff's workbooks. Leave "
                             "unset to apply them to every individual workbook.")
    parser.add_argument("--cathy-cap", type=int, default=None,
                        help="Give Cathy exactly this many professional-pool rows "
                             "for this run instead of the standard weekday schedule. "
                             "Ignored if --no-cathy or --even-split-jasmine-cathy is "
                             "also set.")
    parser.add_argument("--custom-report-name", default=None,
                        help="Staff name for a second, generic Cathy-shaped report: "
                             "with --custom-report-payers, every Insurance row whose "
                             "Payer matches goes to this staff member with their own "
                             "workbook. Must not be Jasmine, CB, Melissa, Cathy, or "
                             "Unable to Bill.")
    parser.add_argument("--custom-report-payers", default="",
                        help="Comma-separated payer terms (case-insensitive substring "
                             "match) for --custom-report-name.")
    parser.add_argument("--custom-report-any-claim-type", action="store_true",
                        help="By default the custom report only claims Professional "
                             "(CMS-1500/UB-04) rows, same as Cathy. Set this to match "
                             "any claim type instead.")
    parser.add_argument("--exclude-payer-by-division-payers", default="",
                        help="Comma-separated funding source (Payer) terms "
                             "(case-insensitive substring match) to remove by "
                             "division. Example: --exclude-payer-by-division-payers "
                             "\"BCBS, Beacon\". Must be combined with "
                             "--exclude-payer-by-division-divisions; both lists "
                             "support multiple values, matched as any payer in any "
                             "division.")
    parser.add_argument("--exclude-payer-by-division-divisions", default="",
                        help="Comma-separated division (GROUPFLD1) terms "
                             "(case-insensitive substring match) the funding-source "
                             "removal above applies to. Example: "
                             "--exclude-payer-by-division-divisions \"Residential, "
                             "Detox, OP Wilton, OP Canaan\". Rows matching neither "
                             "list, or only one of the two lists, are not excluded.")
    args = parser.parse_args()

    if args.workbook_path:
        main(args.workbook_path,
             include_programming=args.include_programming,
             exclude_aetna=args.exclude_aetna,
             cathy_report=args.cathy_report,
             cathy_all_payers=args.cathy_all_payers,
             skip_cathy=args.skip_cathy,
             even_split_jasmine_cathy=args.even_split_jasmine_cathy,
             exclude_payer_terms=parse_terms(args.exclude_payers),
             exclude_service_terms=parse_terms(args.exclude_services),
             exclude_scope=parse_terms(args.exclude_scope),
             cathy_cap_override=args.cathy_cap,
             custom_report_name=args.custom_report_name,
             custom_report_payer_terms=parse_terms(args.custom_report_payers),
             custom_report_professional_only=not args.custom_report_any_claim_type,
             division_exclude_payer_terms=parse_terms(args.exclude_payer_by_division_payers),
             division_exclude_division_terms=parse_terms(args.exclude_payer_by_division_divisions))
