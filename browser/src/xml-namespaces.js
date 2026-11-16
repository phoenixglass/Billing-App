/*
 * Make an .xlsx readable by ExcelJS whatever namespace prefixes its XML uses.
 *
 * In XML, a namespace prefix is only a local nickname: <x:workbook
 * xmlns:x="...spreadsheetml/2006/main"> and <workbook xmlns="...main"> mean
 * exactly the same thing. Excel and openpyxl (the Streamlit app) read both.
 * Files written by some reporting tools (notably ones built on Microsoft's
 * .NET Open XML SDK) use prefixes like that.
 *
 * ExcelJS, though, matches tag names as literal text: it looks for
 * "workbook", never finds "x:workbook", and crashes with errors such as
 * "Cannot read properties of undefined (reading 'sheets')". Likewise it
 * looks for "r:id", "dc:creator", "cp:lastModifiedBy", "x14:..." and so on,
 * spelled exactly that way.
 *
 * So before ExcelJS sees the file, rewrite each XML part so every namespace
 * ExcelJS knows uses the prefix ExcelJS expects (CANONICAL below). Parts that
 * already do are left byte-for-byte as they were, and a file that needs no
 * change at all is passed through untouched.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.XmlNamespaces = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  // Namespace URI -> the prefix ExcelJS's readers expect ("" = no prefix).
  const CANONICAL = new Map([
    ["http://schemas.openxmlformats.org/spreadsheetml/2006/main", ""],
    ["http://schemas.openxmlformats.org/package/2006/relationships", ""],
    ["http://schemas.openxmlformats.org/package/2006/content-types", ""],
    ["http://schemas.openxmlformats.org/officeDocument/2006/extended-properties", ""],
    ["http://schemas.openxmlformats.org/officeDocument/2006/relationships", "r"],
    ["http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes", "vt"],
    ["http://schemas.openxmlformats.org/package/2006/metadata/core-properties", "cp"],
    ["http://purl.org/dc/elements/1.1/", "dc"],
    ["http://purl.org/dc/terms/", "dcterms"],
    ["http://purl.org/dc/dcmitype/", "dcmitype"],
    ["http://www.w3.org/2001/XMLSchema-instance", "xsi"],
    ["http://schemas.openxmlformats.org/markup-compatibility/2006", "mc"],
    ["http://schemas.microsoft.com/office/spreadsheetml/2009/9/main", "x14"],
    ["http://schemas.microsoft.com/office/excel/2006/main", "xm"],
    ["http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing", "xdr"],
    ["http://schemas.openxmlformats.org/drawingml/2006/main", "a"],
  ]);

  // A tag, allowing ">" inside quoted attribute values.
  const TAG_RE = /<(?:[^>"']|"[^"]*"|'[^']*')*>/g;
  const DECL_RE = /\sxmlns(?::([\w.-]+))?\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

  /**
   * Decide which prefixes to rename (Map from -> to, "" meaning no prefix).
   * Returns null when the part needs no change.
   */
  function planRenames(xml) {
    const declared = new Map(); // prefix -> uri, or null if declared with conflicting uris
    let m;
    DECL_RE.lastIndex = 0;
    while ((m = DECL_RE.exec(xml))) {
      const prefix = m[1] || "";
      const uri = m[2] !== undefined ? m[2] : m[3];
      if (!declared.has(prefix)) declared.set(prefix, uri);
      else if (declared.get(prefix) !== uri) declared.set(prefix, null);
    }
    const renames = new Map();
    for (const [prefix, uri] of declared) {
      if (uri === null || !CANONICAL.has(uri)) continue;
      const want = CANONICAL.get(uri);
      if (want !== prefix) renames.set(prefix, want);
    }
    if (!renames.size) return null;
    // Never let a rename collide with a prefix that keeps a different meaning.
    for (const [from, to] of [...renames]) {
      if (declared.has(to) && !renames.has(to) && declared.get(to) !== declared.get(from)) renames.delete(from);
    }
    // Two prefixes may not both become the same one unless they mean the same.
    const target = new Map();
    for (const [from, to] of [...renames]) {
      if (target.has(to) && declared.get(target.get(to)) !== declared.get(from)) renames.delete(from);
      else target.set(to, from);
    }
    return renames.size ? renames : null;
  }

  const withPrefix = (prefix, name) => (prefix ? `${prefix}:${name}` : name);

  function rewriteTag(tag, renames) {
    if (tag.startsWith("<?") || tag.startsWith("<!")) return tag; // declarations, comments, CDATA
    // Element name. An unprefixed element is in the default namespace.
    tag = tag.replace(/^<(\/?)(?:([\w.-]+):)?([\w.-]+)/, (all, slash, prefix, name) => {
      const p = prefix || "";
      return renames.has(p) ? `<${slash}${withPrefix(renames.get(p), name)}` : all;
    });
    // Attributes and declarations, outside quoted values only.
    const pieces = tag.split(/("[^"]*"|'[^']*')/);
    for (let i = 0; i < pieces.length; i += 2) {
      pieces[i] = pieces[i].replace(/(\s)(?:([\w.-]+):)?([\w.-]+)(\s*=)/g, (all, sp, prefix, name, eq) => {
        if (prefix === "xmlns" || (!prefix && name === "xmlns")) {
          const declaredPrefix = prefix ? name : "";
          if (!renames.has(declaredPrefix)) return all;
          const to = renames.get(declaredPrefix);
          return `${sp}${to ? "xmlns:" + to : "xmlns"}${eq}`;
        }
        // An unprefixed attribute is in no namespace: leave it.
        if (!prefix || !renames.has(prefix)) return all;
        return `${sp}${withPrefix(renames.get(prefix), name)}${eq}`;
      });
    }
    tag = pieces.join("");
    // Two declarations of the same namespace can now share a name; keep the first.
    const seen = new Set();
    return tag.replace(/(\s)(xmlns(?::[\w.-]+)?)\s*=\s*("[^"]*"|'[^']*')/g, (all, sp, attr) => {
      if (seen.has(attr)) return "";
      seen.add(attr);
      return all;
    });
  }

  /** Rewrite one XML part. Returns the new text, or null if it needs no change. */
  function normalizeXml(xml) {
    const renames = planRenames(xml);
    if (!renames) return null;
    const out = xml.replace(TAG_RE, (tag) => rewriteTag(tag, renames));
    return out === xml ? null : out;
  }

  const isXmlPart = (name) => /\.(xml|rels)$/i.test(name) && !/^\/?xl\/(theme|media)\//i.test(name);

  // The only worksheet part names ExcelJS reads. It finds worksheets by file
  // name, not through the workbook's relationships the way Excel and
  // openpyxl do, so a sheet stored as e.g. xl/worksheets/sheet.xml (no
  // number) or xl/worksheets/Data.xml is silently skipped. With a second tab
  // in the file, that tab is then read instead and the app reports
  // "Billable Status column not found".
  const SHEET_PART_RE = /^xl\/worksheets\/sheet(\d+)\.xml$/;

  const attr = (tag, name) => {
    const m = tag.match(new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`));
    return m ? (m[1] !== undefined ? m[1] : m[2]) : null;
  };
  const setAttr = (tag, name, value) =>
    tag.replace(new RegExp(`(\\s${name}\\s*=\\s*)(?:"[^"]*"|'[^']*')`), `$1"${value}"`);

  /** Resolve a relationship target against a folder (no trailing slash) to a zip path. */
  function resolvePath(dir, target) {
    const parts = (target.startsWith("/") ? target.slice(1) : `${dir}/${target}`).split("/");
    const out = [];
    for (const p of parts) {
      if (p === "..") out.pop();
      else if (p && p !== ".") out.push(p);
    }
    return out.join("/");
  }

  /** Zip entry for a part name; OPC part names are case-insensitive. */
  function findEntry(zip, path) {
    const exact = zip.file(path);
    if (exact) return exact;
    const lower = path.toLowerCase();
    return Object.values(zip.files).find((f) => !f.dir && f.name.toLowerCase() === lower) || null;
  }

  /**
   * Move every worksheet part ExcelJS wouldn't find to xl/worksheets/sheetN.xml,
   * updating the workbook's relationships, [Content_Types].xml and the sheet's
   * own relationships. Expects namespace prefixes already normalized.
   * Returns the list of parts moved, as "old -> new".
   */
  async function renameSheetParts(zip) {
    const relsName = "xl/_rels/workbook.xml.rels";
    const relsEntry = findEntry(zip, relsName);
    if (!relsEntry) return [];
    const rels = await relsEntry.async("string");

    const used = new Set();
    for (const name of Object.keys(zip.files)) {
      const m = name.match(SHEET_PART_RE);
      if (m) used.add(+m[1]);
    }
    let next = 1;
    const freshName = () => {
      while (used.has(next)) next++;
      used.add(next);
      return `xl/worksheets/sheet${next}.xml`;
    };

    const moves = [];
    const newRels = rels.replace(/<Relationship\b[^>]*>/g, (tag) => {
      const type = attr(tag, "Type") || "";
      const target = attr(tag, "Target");
      if (!/\/worksheet$/.test(type) || !target || attr(tag, "TargetMode") === "External") return tag;
      const oldPath = resolvePath("xl", target);
      const entry = findEntry(zip, oldPath);
      if (!entry || SHEET_PART_RE.test(entry.name)) return tag;
      const newPath = freshName();
      moves.push({ from: entry.name, to: newPath });
      return setAttr(tag, "Target", newPath.slice("xl/".length));
    });
    if (!moves.length) return [];

    zip.file(relsEntry.name, newRels);
    for (const { from, to } of moves) {
      zip.file(to, await zip.file(from).async("uint8array"));
      zip.remove(from);

      // The sheet's own relationships (drawings, hyperlinks, ...): move them
      // too, re-pointing relative targets from the old folder to the new one.
      const fromDir = from.slice(0, from.lastIndexOf("/"));
      const fromBase = from.slice(from.lastIndexOf("/") + 1);
      const sheetRels = findEntry(zip, `${fromDir}/_rels/${fromBase}.rels`);
      if (sheetRels) {
        let text = await sheetRels.async("string");
        text = text.replace(/<Relationship\b[^>]*>/g, (tag) => {
          const t = attr(tag, "Target");
          if (!t || t.startsWith("/") || attr(tag, "TargetMode") === "External") return tag;
          const abs = resolvePath(fromDir, t);
          return setAttr(tag, "Target", abs.startsWith("xl/") ? "../" + abs.slice(3) : "/" + abs);
        });
        zip.remove(sheetRels.name);
        zip.file(`xl/worksheets/_rels/${to.slice("xl/worksheets/".length)}.rels`, text);
      }
    }

    const ctEntry = findEntry(zip, "[Content_Types].xml");
    if (ctEntry) {
      let ct = await ctEntry.async("string");
      for (const { from, to } of moves) {
        ct = ct.replace(/<Override\b[^>]*>/g, (tag) =>
          (attr(tag, "PartName") || "").toLowerCase() === "/" + from.toLowerCase()
            ? setAttr(tag, "PartName", "/" + to) : tag);
      }
      zip.file(ctEntry.name, ct);
    }
    return moves.map(({ from, to }) => `${from} -> ${to}`);
  }

  /**
   * Normalize every XML part of an .xlsx, then move worksheet parts ExcelJS
   * can't find to names it can. `data` is an ArrayBuffer/Buffer/
   * Uint8Array. Returns { data, changed }: `data` is the original object when
   * nothing needed rewriting, and `changed` lists the parts that were.
   */
  async function normalizeXlsx(JSZip, data) {
    let zip;
    try {
      zip = await JSZip.loadAsync(data);
    } catch (e) {
      return { data, changed: [] }; // not a zip: let ExcelJS report it
    }
    const changed = [];
    for (const entry of Object.values(zip.files)) {
      if (entry.dir || !isXmlPart(entry.name)) continue;
      const bytes = await entry.async("uint8array");
      // UTF-16 parts (BOM FF FE / FE FF) are left alone.
      if (bytes.length >= 2 && ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff))) continue;
      const text = new TextDecoder("utf-8").decode(bytes);
      if (!/\sxmlns/.test(text)) continue;
      const out = normalizeXml(text);
      if (out !== null) {
        zip.file(entry.name, out);
        changed.push(entry.name);
      }
    }
    changed.push(...(await renameSheetParts(zip)));
    if (!changed.length) return { data, changed };
    const rebuilt = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return { data: rebuilt, changed };
  }

  return { normalizeXml, normalizeXlsx, renameSheetParts, CANONICAL };
});
