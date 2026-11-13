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

  /**
   * Normalize every XML part of an .xlsx. `data` is an ArrayBuffer/Buffer/
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
    if (!changed.length) return { data, changed };
    const rebuilt = await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
    return { data: rebuilt, changed };
  }

  return { normalizeXml, normalizeXlsx, CANONICAL };
});
