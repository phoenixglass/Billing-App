# Billing App — browser version (offline, single file)

`dist/Billing App.html` is the whole Billing App in one file. Double-click it
and it opens in Chrome or Edge. It has the same options and produces the same
workbooks as the Streamlit app, but:

- **Nothing is uploaded.** The spreadsheet is read and processed by the browser
  on your own computer. The page has no network access at all (its Content
  Security Policy blocks every connection), so it also works offline.
- **Nothing to install.** No Python, no `.exe`, no admin rights.
- The finished reports (Jasmine, Cathy, CB, any custom report, and Masters) are
  saved to your Downloads folder.

## Using it

1. Download `dist/Billing App.html` (on GitHub: open the file, then
   **Download raw file**) and save it somewhere handy, like your Desktop.
2. Double-click it. It opens in your browser.
3. Choose the day's export (`.xlsx`, with the date in the filename as
   MMDDYYYY), tick any options, and click **Process file**.
4. Click each download, or **Download all**. The browser may ask once whether
   the page can download several files: allow it.

## How it's built

| File | What it is |
|---|---|
| `src/billing-core.js` | The billing rules: a port of `billing_rules.py` plus the report-building part of `app.py`'s `process_workbook`. |
| `src/workbook-io.js` | Reads the export and writes the output workbooks, using [ExcelJS](https://github.com/exceljs/exceljs) 4.4.0. |
| `src/xml-namespaces.js` | Runs before ExcelJS reads the export. Some exporters write the workbook's XML with namespace prefixes (`<x:workbook>`); Excel and openpyxl accept that, but ExcelJS fails with errors like *Cannot read properties of undefined (reading 'sheets')*. This rewrites those parts (using [JSZip](https://stuk.github.io/jszip/) 3.10.1) to the spelling ExcelJS expects. |
| `src/app.js`, `src/index.template.html` | The page. |
| `build.js` | Inlines ExcelJS, JSZip and the sources into `dist/Billing App.html`. |

```bash
cd browser
npm install
npm run build
```

**The rules exist in two places**: `billing_rules.py` (Python app) and
`src/billing-core.js` (this). Any rule change must be made in both, then
checked with the comparison test below, then rebuilt and committed.

## Tests

`test/compare_with_python.py` runs the same input workbooks through the Python
app's real `process_workbook` and through this version's code (in Node, via
`test/run.js`), then compares every output workbook: file names, sheets, every
cell value, the Status dropdown, column widths, and the summary numbers. It
covers the demo files (if given) and random reports built from the real
export's columns and category values, on every weekday, with 14 combinations of
the per-run options. The first `--prefixed-seeds` inputs (default 2) are also
re-saved with namespace-prefixed XML and run again, to check the browser
version reads those files exactly as the Python app does. No patient data is
involved.

`test/namespaces_test.js` checks the XML rewriting on small hand-written
cases (`node browser/test/namespaces_test.js`).

```bash
# from the repo root, with the app's requirements installed and npm install done
python browser/test/compare_with_python.py --seeds 8
```

`test/browser_test.js` drives the built page in Chromium with Playwright the
way a person would (choose a file, tick options, click Process, download) and
reports any network request the page attempted:

```bash
node browser/test/browser_test.js "path/to/Unbilled Revenue 11102026.xlsx" out '{"phpToJasmine": true}'
```

## If a file won't open

If the page says a file *couldn't be read as an Excel workbook*, open it in
Excel, choose **File > Save As > Excel Workbook (.xlsx)**, and process the
saved copy. Excel rewrites the file in its own standard form.

## Differences from the Streamlit app

- Completely empty rows in the export are skipped instead of being carried into
  the Masters as blank "Unable to Bill" rows.
- In the Masters, each row keeps its own cell formatting when rows are
  reordered, and column widths shift right with the new Staff/Status column.
  The Streamlit app leaves formatting and widths where they were.
- In the individual reports, dates keep the export's date format instead of
  showing as `yyyy-mm-dd h:mm:ss`.
- Cells containing formulas are read by their calculated value.

None of these change who gets which row.
