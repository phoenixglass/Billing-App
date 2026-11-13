/*
 * Reading the uploaded export and writing the output workbooks, with ExcelJS.
 * Runs in the page and in the Node test harness (same code in both).
 *
 * Output layout matches app.py / billing_rules.finalize_workbook:
 * - Masters: the uploaded workbook with Staff/Status in column A, rows
 *   reordered (Jasmine/Cathy pool first, by Client), header bold, a filter on
 *   the header row, other tabs kept, and an "Invalid" tab if any rows were
 *   "Invalid for Billing".
 * - Each person's report: "Sheet1" with their rows; Jasmine/Cathy/custom
 *   reports get Status and Comments inserted at E and F with a dropdown fed
 *   from "Sheet2"; CB's has no Status/Comments. Header bold, columns sized
 *   to fit, filter on the header row.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(require("./billing-core.js"), require("./xml-namespaces.js"), require("jszip"));
  } else root.WorkbookIO = factory(root.BillingCore, root.XmlNamespaces, root.JSZip);
})(typeof self !== "undefined" ? self : this, function (Core, Ns, JSZip) {
  "use strict";

  const DATE_FMT = "mm/dd/yyyy";

  const isBlank = (v) => {
    const p = Core.plain(v);
    return p === null || p === "";
  };
  const clone = (o) => (o ? JSON.parse(JSON.stringify(o)) : null);

  function colLetter(n) {
    let s = "";
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  /** Load an .xlsx (ArrayBuffer/Buffer) and turn its active sheet into a table. */
  async function readWorkbook(ExcelJS, data) {
    // Some exporters write the XML with namespace prefixes (<x:workbook>),
    // which ExcelJS can't read; rewrite those parts first. See xml-namespaces.js.
    const fixed = await Ns.normalizeXlsx(JSZip, data);
    const wb = new ExcelJS.Workbook();
    try {
      await wb.xlsx.load(fixed.data);
    } catch (e) {
      throw new Error("This file couldn't be read as an Excel workbook (" + (e.message || e) + "). " +
        "Try opening it in Excel, choosing File > Save As > Excel Workbook (.xlsx), and processing the saved copy.");
    }
    const sheets = wb.worksheets;
    if (!sheets.length) throw new Error("The file has no worksheets.");
    let active = (wb.views && wb.views[0] && wb.views[0].activeTab) || 0;
    if (!(active >= 0 && active < sheets.length)) active = 0;
    const ws = sheets[active];

    let colCount = 0;
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (c, col) => { if (!isBlank(c.value) && col > colCount) colCount = col; });
    });
    const grab = (row) => {
      const out = [];
      for (let c = 1; c <= colCount; c++) {
        const cell = row.getCell(c);
        out.push({ v: cell.value === undefined ? null : cell.value, s: clone(cell.style) });
      }
      return out;
    };
    const headers = grab(ws.getRow(1));
    const rows = [];
    for (let r = 2; r <= ws.rowCount; r++) {
      const cells = grab(ws.getRow(r));
      if (cells.some((c) => !isBlank(c.v))) rows.push(cells); // skip completely empty rows
    }
    const widths = [];
    for (let c = 1; c <= colCount; c++) widths.push(ws.getColumn(c).width);
    return { workbook: wb, sheet: ws, activeIndex: active, table: { headers, rows }, widths };
  }

  function writeValue(target, cell, keepStyle) {
    const v = cell ? cell.v : null;
    target.value = v === undefined ? null : v;
    if (keepStyle && cell && cell.s) target.style = clone(cell.s);
    const nf = cell && cell.s && cell.s.numFmt;
    if (nf) target.numFmt = nf;
    else if (Core.plain(v) instanceof Date) target.numFmt = DATE_FMT;
  }

  /** Copy a worksheet we don't touch (e.g. a notes tab) into another workbook. */
  function copySheet(src, dst) {
    const ws = dst.addWorksheet(src.name, { views: clone(src.views) || [], properties: clone(src.properties) || {} });
    src.columns && src.columns.forEach((col, i) => { if (col && col.width) ws.getColumn(i + 1).width = col.width; });
    src.eachRow({ includeEmpty: true }, (row, r) => {
      if (row.height) ws.getRow(r).height = row.height;
      row.eachCell({ includeEmpty: true }, (c, col) => {
        const t = ws.getRow(r).getCell(col);
        t.value = c.value === undefined ? null : c.value;
        if (c.style) t.style = clone(c.style);
      });
    });
    const merges = (src.model && src.model.merges) || [];
    merges.forEach((m) => { try { ws.mergeCells(m); } catch (e) { /* ignore odd merges */ } });
    return ws;
  }

  function fitWidths(ws, rows) {
    const n = rows.reduce((m, r) => Math.max(m, r.length), 0);
    for (let c = 0; c < n; c++) {
      let max = 0;
      for (const r of rows) {
        const v = r[c] ? Core.plain(r[c].v) : null;
        if (v === null) continue;
        max = Math.max(max, Core.pyStr(v).length);
      }
      // ExcelJS drops a width of exactly 9 (its default) from the file, which
      // would leave Excel's own default (~8.4); 9.01 looks the same and sticks.
      const w = Math.min(max + 2, 60);
      ws.getColumn(c + 1).width = w === 9 ? 9.01 : w;
    }
  }

  /** Build every output workbook. Returns [{ filename, buffer }]. */
  async function buildOutputs(ExcelJS, loaded, result) {
    const { table } = loaded;
    const outputs = [];

    for (const rep of result.reports) {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("Sheet1");
      let header = table.headers.map((h) => ({ v: h.v }));
      let rows = rep.rows.map((r) => r.map((c) => ({ v: c.v, s: c.s && c.s.numFmt ? { numFmt: c.s.numFmt } : null })));
      if (rep.statusColumns) {
        header.splice(4, 0, { v: "Status" }, { v: "Comments" });
        rows = rows.map((r) => { const x = r.slice(); x.splice(4, 0, { v: null }, { v: null }); return x; });
      }
      [header].concat(rows).forEach((r, i) => {
        const row = ws.getRow(i + 1);
        r.forEach((c, j) => writeValue(row.getCell(j + 1), c, false));
      });
      ws.getRow(1).eachCell({ includeEmpty: true }, (c) => { c.font = { bold: true }; });
      if (rep.statusColumns) {
        const list = wb.addWorksheet("Sheet2");
        result.statusItems.forEach((item, i) => { list.getCell(`A${i + 1}`).value = item; });
        const formula = `Sheet2!$A$1:$A$${result.statusItems.length}`;
        for (let r = 2; r <= rows.length + 1; r++) {
          ws.getCell(`E${r}`).dataValidation = { type: "list", allowBlank: true, formulae: [formula] };
        }
      }
      fitWidths(ws, [header].concat(rows));
      ws.autoFilter = `A1:${colLetter(header.length)}${rows.length + 1}`;
      outputs.push({ filename: rep.filename, buffer: await wb.xlsx.writeBuffer() });
    }

    // Masters: the uploaded workbook, processed sheet swapped in, Invalid tab added.
    const wb = new ExcelJS.Workbook();
    loaded.workbook.worksheets.forEach((src, i) => {
      if (i !== loaded.activeIndex) { copySheet(src, wb); return; }
      const ws = wb.addWorksheet(src.name, { views: clone(src.views) || [] });
      // Column widths move right one place with the new Staff/Status column.
      const shift = table.headers.length - loaded.widths.length;
      loaded.widths.forEach((w, c) => { if (w) ws.getColumn(c + 1 + shift).width = w; });
      [table.headers].concat(table.rows).forEach((r, ri) => {
        const row = ws.getRow(ri + 1);
        r.forEach((c, j) => writeValue(row.getCell(j + 1), c, true));
      });
      ws.getRow(1).eachCell({ includeEmpty: true }, (c) => { c.font = Object.assign({}, c.font || {}, { bold: true }); });
      ws.autoFilter = `A1:${colLetter(table.headers.length)}${table.rows.length + 1}`;
    });
    if (result.invalidRows.length) {
      const inv = wb.addWorksheet("Invalid");
      [result.invalidHeaders].concat(result.invalidRows).forEach((r, ri) => {
        const row = inv.getRow(ri + 1);
        r.forEach((c, j) => writeValue(row.getCell(j + 1), { v: c.v, s: c.s && c.s.numFmt ? { numFmt: c.s.numFmt } : null }, false));
      });
    }
    wb.views = [{ activeTab: loaded.activeIndex, firstSheet: 0, visibility: "visible" }];
    outputs.push({ filename: result.mastersFilename, buffer: await wb.xlsx.writeBuffer() });
    return outputs;
  }

  /** Read, process and write in one go. */
  async function run(ExcelJS, data, filename, opts) {
    const loaded = await readWorkbook(ExcelJS, data);
    const result = Core.processTable(loaded.table, filename, opts);
    const files = await buildOutputs(ExcelJS, loaded, result);
    return { result, files };
  }

  return { readWorkbook, buildOutputs, run };
});
