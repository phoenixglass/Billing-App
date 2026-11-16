#!/usr/bin/env node
/*
 * Build the single-file app: browser/dist/Billing App.html
 *
 * Inlines ExcelJS, JSZip and the source files into the HTML template, so the
 * result is one file that works offline with nothing to install. Run
 * `npm install` once, then `npm run build`.
 */
const fs = require("fs");
const path = require("path");

const root = __dirname;
const read = (p) => fs.readFileSync(path.join(root, p), "utf8");
// A literal "</script" inside inlined code would end the <script> tag early.
const safe = (js) => js.replace(/<\/script/gi, "<\\/script");

const parts = {
  "/*EXCELJS*/": read("node_modules/exceljs/dist/exceljs.min.js"),
  "/*JSZIP*/": read("node_modules/jszip/dist/jszip.min.js"),
  "/*NS*/": read("src/xml-namespaces.js"),
  "/*CORE*/": read("src/billing-core.js"),
  "/*IO*/": read("src/workbook-io.js"),
  "/*APP*/": read("src/app.js"),
};
let html = read("src/index.template.html");
for (const [marker, code] of Object.entries(parts)) {
  if (!html.includes(marker)) throw new Error(`template is missing ${marker}`);
  html = html.split(marker).join(safe(code));
}
const outDir = path.join(root, "dist");
fs.mkdirSync(outDir, { recursive: true });
const out = path.join(outDir, "Billing App.html");
fs.writeFileSync(out, html);
console.log(`wrote ${out} (${(html.length / 1024).toFixed(0)} KB)`);
