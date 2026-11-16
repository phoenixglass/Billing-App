#!/usr/bin/env node
/* Unit checks for src/xml-namespaces.js.  node test/namespaces_test.js */
const assert = require("assert");
const { normalizeXml } = require("../src/xml-namespaces.js");
const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC = "http://purl.org/dc/elements/1.1/";
const cases = [
  ["already canonical: untouched",
    `<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="A" r:id="rId1"/></sheets></workbook>`, null],
  ["main namespace prefixed",
    `<?xml version="1.0"?><x:workbook xmlns:x="${MAIN}" xmlns:r="${R}"><x:sheets><x:sheet name="A" sheetId="1" r:id="rId1"/></x:sheets></x:workbook>`,
    `<?xml version="1.0"?><workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="A" sheetId="1" r:id="rId1"/></sheets></workbook>`],
  ["relationships prefix renamed to r",
    `<workbook xmlns="${MAIN}" xmlns:rel="${R}"><sheets><sheet name="A" rel:id="rId1"/></sheets></workbook>`,
    `<workbook xmlns="${MAIN}" xmlns:r="${R}"><sheets><sheet name="A" r:id="rId1"/></sheets></workbook>`],
  ["text and quoted values that look like prefixes are left alone",
    `<x:sst xmlns:x="${MAIN}"><x:si><x:t>a x:b="c" &lt;x:d&gt;</x:t></x:si><x:si><x:t xml:space="preserve" v='x:y="1" >'>q</x:t></x:si></x:sst>`,
    `<sst xmlns="${MAIN}"><si><t>a x:b="c" &lt;x:d&gt;</t></si><si><t xml:space="preserve" v='x:y="1" >'>q</t></si></sst>`],
  ["default namespace moved to its expected prefix",
    `<coreProperties xmlns="${CP}" xmlns:d="${DC}"><d:creator>Me</d:creator><lastModifiedBy>Me</lastModifiedBy></coreProperties>`,
    `<cp:coreProperties xmlns:cp="${CP}" xmlns:dc="${DC}"><dc:creator>Me</dc:creator><cp:lastModifiedBy>Me</cp:lastModifiedBy></cp:coreProperties>`],
  ["both prefixed and default for the same namespace: one declaration left",
    `<x:worksheet xmlns="${MAIN}" xmlns:x="${MAIN}"><sheetData><x:row r="1"/></sheetData></x:worksheet>`,
    `<worksheet xmlns="${MAIN}"><sheetData><row r="1"/></sheetData></worksheet>`],
  ["a different default namespace blocks the rename",
    `<x:workbook xmlns="urn:other" xmlns:x="${MAIN}"><x:sheets/></x:workbook>`, null],
  ["unknown namespaces untouched",
    `<workbook xmlns="${MAIN}" xmlns:foo="urn:foo"><foo:bar foo:a="1"/></workbook>`, null],
];
let failed = 0;
for (const [name, input, want] of cases) {
  const got = normalizeXml(input);
  try { assert.strictEqual(got, want); console.log("ok      ", name); }
  catch (e) { failed++; console.log("FAIL    ", name, "\n  got: ", got, "\n  want:", want); }
}
console.log(failed ? `${failed} failed` : "all passed");
process.exit(failed ? 1 : 0);
