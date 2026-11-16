#!/usr/bin/env node
/*
 * Run the browser version's processing in Node on one file.
 *   node test/run.js INPUT.xlsx OUTDIR '{"excludeAetna":true}'
 * Writes the output workbooks to OUTDIR plus summary.json. Used by
 * test/compare_with_python.py; the page itself runs the same modules.
 */
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const IO = require("../src/workbook-io.js");

async function main() {
  const [input, outdir, optsJson, nameOverride] = process.argv.slice(2);
  const opts = optsJson ? JSON.parse(optsJson) : {};
  fs.mkdirSync(outdir, { recursive: true });
  const filename = nameOverride || path.basename(input);
  let summary;
  try {
    const { result, files } = await IO.run(ExcelJS, fs.readFileSync(input), filename, opts);
    for (const f of files) fs.writeFileSync(path.join(outdir, f.filename), Buffer.from(f.buffer));
    summary = {
      ok: true, files: files.map((f) => f.filename), invalidCount: result.invalidCount,
      dateToken: result.dateToken, wmCount: result.wmCount, splitDay: result.splitDay,
    };
  } catch (e) {
    summary = { ok: false, error: e.message };
  }
  fs.writeFileSync(path.join(outdir, "summary.json"), JSON.stringify(summary, null, 2));
  console.log(JSON.stringify(summary));
}
main();
