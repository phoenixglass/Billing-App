/*
 * Billing rules for the browser version of the Billing App.
 *
 * This is a line-by-line port of billing_rules.py plus the parts of app.py
 * that decide which rows go into which report (process_workbook). It has no
 * Excel or browser dependencies: it works on a plain table of cells, so the
 * same file runs in the page and in the Node test harness, which checks it
 * row-by-row against the Python app (see test/compare_with_python.py).
 *
 * If you change a rule here, make the same change in billing_rules.py (or
 * vice versa) and re-run the comparison; the two must agree.
 *
 * A table is { headers: [value...], rows: [[cell...]] } where each cell is
 * { v: value, s: style } and value is whatever ExcelJS read (string, number,
 * Date, null, ...). Styles are carried along untouched so dates keep their
 * formats.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.BillingCore = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const DRUG_SCREEN_KEYWORDS = ["drug screen", "utox", "urine tox", "drug test", "uds"];
  const PHP_KEYWORDS = ["partial hospitalization", "php"];
  const ECARE_KEYWORDS = ["e-care", "e care", "ecare", "extended care"];
  // 09/23/2026 is a known Day A; every calendar day alternates A, B, A, B.
  const SPLIT_DAY_A_ANCHOR_UTC = Date.UTC(2026, 8, 23);
  const SPLIT_DAYS = ["A", "B"];
  const RESERVED_STAFF_NAMES = ["Jasmine", "CB", "Melissa", "Cathy", "Unable to Bill"];
  const STATUS_ITEMS = ["Billed", "Unable to Bill", "Contractual Adj", "Incomplete Billings",
    "Utox Batch", "Inclusive Services"];

  // ---- value helpers -------------------------------------------------------

  /** Unwrap an ExcelJS cell value (formula, rich text, hyperlink) to a plain value. */
  function plain(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return v;
    if (typeof v === "object") {
      if ("result" in v) return plain(v.result);
      if (Array.isArray(v.richText)) return v.richText.map((r) => r.text).join("");
      if ("text" in v) return plain(v.text);
      if ("error" in v) return v.error;
      return null;
    }
    return v;
  }

  function pad(n, w) { return String(n).padStart(w || 2, "0"); }

  /** Python's str() for the kinds of values a cell can hold. */
  function pyStr(v) {
    v = plain(v);
    if (v === null) return "None";
    if (v instanceof Date) {
      return `${v.getUTCFullYear()}-${pad(v.getUTCMonth() + 1)}-${pad(v.getUTCDate())} ` +
        `${pad(v.getUTCHours())}:${pad(v.getUTCMinutes())}:${pad(v.getUTCSeconds())}`;
    }
    if (typeof v === "boolean") return v ? "True" : "False";
    return String(v);
  }

  /** Python's str(value or ""). */
  function text(v) {
    v = plain(v);
    if (v === null || v === "" || v === 0 || v === false) return "";
    return pyStr(v);
  }

  const lower = (s) => (s || "").toLowerCase();
  const containsAny = (s, words) => words.some((w) => s.includes(w));

  // ---- classifiers (billing_rules.py) --------------------------------------

  const isEcare = (service) => containsAny(lower(service), ECARE_KEYWORDS);
  const isDrugScreen = (service) => containsAny(lower(service), DRUG_SCREEN_KEYWORDS);
  function isPhpService(service) {
    const s = lower(service);
    if (isEcare(s)) return false; // "E-care PHP Housing" is e-care, not PHP
    return containsAny(s, PHP_KEYWORDS);
  }
  const isWithdrawalManagementService = (service) => lower(service).includes("withdrawal management");
  function isProgrammingService(service) {
    const s = lower(service);
    return s.includes("detox") || s.includes("residential");
  }
  const isIopService = (service) => lower(service).includes("iop");
  const isDayTreatmentService = (service) => lower(service).includes("day treatment");
  const billsLikeIop = (service) => isIopService(service) || isDayTreatmentService(service);
  const isProfessionalClaimType = (claimType) => (claimType || "").trim().toUpperCase() === "CMS-1500";
  const isAetnaPayer = (payer) => lower(payer).includes("aetna");
  const isAnthemPayer = (payer) => lower(payer).includes("anthem");
  const isWmProgramLevel = (v) => (v === null || v === undefined ? "" : text(v)).toUpperCase().includes("WM");
  function isBcbAnthemCtPhpResDetox(payer, service) {
    return lower(payer).includes("bcb anthem ct") &&
      ["partial hospitalization", "residential", "detox"].some((s) => lower(service).includes(s));
  }

  function parseTerms(raw) {
    if (!raw) return [];
    return raw.split(",").map((t) => t.trim()).filter(Boolean);
  }
  function matchesAnyTerm(str, terms) {
    if (!terms || !terms.length) return false;
    const hay = lower(str);
    return terms.some((t) => hay.includes(t.toLowerCase()));
  }
  function payerExcludedByDivision(payer, division, payerTerms, divisionTerms) {
    if (!payerTerms.length || !divisionTerms.length) return false;
    return matchesAnyTerm(payer, payerTerms) && matchesAnyTerm(division, divisionTerms);
  }

  function isNonBillableServiceForWeekday(service, weekday, opts) {
    opts = opts || {};
    const s = lower(service).trim();
    if (opts.selfPay) return false;
    if (isEcare(s)) return weekday !== 1; // Tuesday only, whatever the claim type
    if (opts.isProfessional) return false;
    if (billsLikeIop(s)) return false;
    if (isProgrammingService(s)) {
      if (opts.includeProgramming) return false;
      return ![1, 3, 4, 5, 6].includes(weekday);
    }
    return true;
  }

  class ReportExclusions {
    constructor(o) {
      o = o || {};
      this.excludeOptum = !!o.excludeOptum;
      this.excludeBcbAnthemCt = !!o.excludeBcbAnthemCt;
      this.excludeAnthemCathyJasmine = !!o.excludeAnthemCathyJasmine;
      this.excludeAetna = !!o.excludeAetna;
      this.excludeDetoxResidential = !!o.excludeDetoxResidential;
      this.payerTerms = o.payerTerms || [];
      this.serviceTerms = o.serviceTerms || [];
      this.scope = o.scope || [];
      this.divisionPayerTerms = o.divisionPayerTerms || [];
      this.divisionTerms = o.divisionTerms || [];
    }
    excludes(staff, payer, service, division) {
      payer = payer || "";
      service = service || "";
      if (this.excludeOptum && isDrugScreen(service) && lower(payer).includes("optum")) return true;
      if (this.excludeBcbAnthemCt && isBcbAnthemCtPhpResDetox(payer, service)) return true;
      if (this.excludeAnthemCathyJasmine && (staff === "Cathy" || staff === "Jasmine") && isAnthemPayer(payer)) return true;
      if (this.excludeAetna && isAetnaPayer(payer)) return true;
      if (this.excludeDetoxResidential && isProgrammingService(service)) return true;
      if (payerExcludedByDivision(payer, division || "", this.divisionPayerTerms, this.divisionTerms)) return true;
      if (!this.scope.length || this.scope.includes(staff)) {
        if (matchesAnyTerm(payer, this.payerTerms)) return true;
        if (matchesAnyTerm(service, this.serviceTerms)) return true;
      }
      return false;
    }
    splitExclusions(payer, service, division) {
      return [this.excludes("Jasmine", payer, service, division), this.excludes("Cathy", payer, service, division)];
    }
  }

  function validateCustomReportName(name) {
    if (name && RESERVED_STAFF_NAMES.map((n) => n.toLowerCase()).includes(name.trim().toLowerCase())) {
      throw new Error(`'${name}' is already a reserved staff name (${RESERVED_STAFF_NAMES.join(", ")}); ` +
        "choose a different name for the custom report.");
    }
  }

  // ---- dates ---------------------------------------------------------------

  function extractDateFromFilename(filename) {
    const m = /(\d{8})/.exec(filename || "");
    return m ? m[1] : null;
  }

  function getFilenamePrefix(filename) {
    const m = /(\d{8})/.exec(filename || "");
    if (!m) return "";
    let prefix = filename.slice(0, m.index + 8);
    if (filename.slice(m.index + 8).startsWith(" - ")) prefix += " - ";
    return prefix;
  }

  /** MMDDYYYY -> UTC midnight ms, or null if it isn't a real date (like strptime). */
  function tokenToUtc(token) {
    if (!token || !/^\d{8}$/.test(token)) return null;
    const mm = +token.slice(0, 2), dd = +token.slice(2, 4), yyyy = +token.slice(4);
    const t = Date.UTC(yyyy, mm - 1, dd);
    const d = new Date(t);
    if (d.getUTCFullYear() !== yyyy || d.getUTCMonth() !== mm - 1 || d.getUTCDate() !== dd) return null;
    return t;
  }

  function todayUtc() {
    const n = new Date();
    return Date.UTC(n.getFullYear(), n.getMonth(), n.getDate());
  }

  /** Weekday with Monday=0 ... Sunday=6, like Python. */
  function parseWeekdayFromToken(token) {
    let t = tokenToUtc(token), fallback = false;
    if (t === null) { t = todayUtc(); fallback = true; }
    return { weekday: (new Date(t).getUTCDay() + 6) % 7, fallback };
  }

  function splitDayForDateToken(token) {
    let t = tokenToUtc(token), fallback = false;
    if (t === null) { t = todayUtc(); fallback = true; }
    const days = Math.round((t - SPLIT_DAY_A_ANCHOR_UTC) / 86400000);
    return { day: SPLIT_DAYS[((days % 2) + 2) % 2], fallback };
  }

  // ---- the pipeline --------------------------------------------------------

  const cell = (v, s) => ({ v: v === undefined ? null : v, s: s || null });
  const val = (row, col) => (row[col - 1] ? row[col - 1].v : null); // 1-based column, like openpyxl

  /** step_1_extract_invalid: add Staff/Status, pull "Invalid for Billing" rows out. */
  function extractInvalid(table) {
    if (plain(table.headers[0] ? table.headers[0].v : null) !== "Staff/Status") {
      table.headers.unshift(cell("Staff/Status"));
      table.rows.forEach((r) => r.unshift(cell(null)));
    }
    const billableCol = table.headers.findIndex((h) => plain(h.v) === "Billable Status") + 1;
    if (!billableCol) throw new Error("Billable Status column not found");
    const invalid = [], keep = [];
    for (const r of table.rows) (plain(val(r, billableCol)) === "Invalid for Billing" ? invalid : keep).push(r);
    table.rows = keep;
    return invalid;
  }

  /** Find columns by header; like the Python loop, a later duplicate header wins. */
  function findColumns(headers) {
    const names = {
      GROUPFLD1: "groupFld1", GROUPFLD2: "group", Service: "service", Payer: "payer",
      "Billing Provider": "billingProvider", "Program Level": "programLevel",
      Client: "client", "Claim Type": "claimType",
    };
    const cols = {};
    headers.forEach((h, i) => {
      const key = names[plain(h.v)];
      if (key) cols[key] = i + 1;
    });
    return cols;
  }

  /** assign_staff: fills column 1 and reorders rows (shared pool first, by Client). */
  function assignStaff(table, dateToken, o) {
    o = o || {};
    const cols = findColumns(table.headers);
    if (!["group", "service", "payer", "claimType", "client"].every((k) => cols[k])) {
      throw new Error("Missing required columns: GROUPFLD2, Service, Payer, Claim Type, or Client");
    }
    const weekday = parseWeekdayFromToken(dateToken).weekday;
    let splitDay = o.splitDay;
    if (splitDay === null || splitDay === undefined || splitDay === "") {
      splitDay = splitDayForDateToken(dateToken).day;
    } else {
      splitDay = String(splitDay).trim().toUpperCase();
      if (!SPLIT_DAYS.includes(splitDay)) throw new Error(`split_day must be 'A' or 'B', not '${o.splitDay}'`);
    }
    const exclusions = o.exclusions || new ReportExclusions();
    const customActive = !!(o.customReportName && o.customReportPayerTerms && o.customReportPayerTerms.length);
    const professionalOnly = o.customReportProfessionalOnly !== false;

    const fixed = new Map(), pool = [], others = [];
    for (const row of table.rows) {
      const group = text(val(row, cols.group)).trim();
      const service = text(val(row, cols.service));
      const payer = text(val(row, cols.payer)).toLowerCase();
      const claimType = text(val(row, cols.claimType));

      if (group === "Self Pay") { fixed.set(row, "CB"); others.push(row); continue; }

      let staff = null;
      if ((cols.programLevel && isWmProgramLevel(val(row, cols.programLevel))) ||
          isWithdrawalManagementService(service)) staff = "Melissa";
      if (!staff && cols.billingProvider && cols.groupFld1) {
        const bp = text(val(row, cols.billingProvider)).trim();
        const gf1 = text(val(row, cols.groupFld1)).trim();
        if (bp === "O'Flynn, Karen" && (gf1 === "OP Chappaqua" || gf1 === "OP NYC")) staff = "Unable to Bill";
      }
      if (!staff) {
        const sl = service.toLowerCase();
        if ((sl.includes("detox") || sl.includes("residential")) &&
            (payer.includes("aetna") || payer.includes("humana")) && !isDrugScreen(service)) staff = "Melissa";
      }
      if (!staff && isPhpService(service)) {
        staff = o.phpToJasmine && group === "Insurance" ? "Jasmine" : "Melissa";
      }
      if (staff) { fixed.set(row, staff); others.push(row); continue; }

      if (group !== "Insurance") { fixed.set(row, "Unable to Bill"); others.push(row); continue; }

      if (customActive && (!professionalOnly || isProfessionalClaimType(claimType)) &&
          matchesAnyTerm(payer, o.customReportPayerTerms)) {
        fixed.set(row, o.customReportName); others.push(row); continue;
      }

      if (isNonBillableServiceForWeekday(service, weekday, {
        isProfessional: isProfessionalClaimType(claimType), includeProgramming: !!o.includeProgramming,
      })) { fixed.set(row, "Unable to Bill"); others.push(row); continue; }

      const client = text(val(row, cols.client)).trim();
      const division = cols.groupFld1 ? text(val(row, cols.groupFld1)) : "";
      const [exJ, exC] = exclusions.splitExclusions(text(val(row, cols.payer)), service, division);
      pool.push({ row, key: client.toLowerCase(), exJ, exC });
    }

    // Stable sort by lower-cased Client, comparing like Python strings.
    pool.sort((a, b) => (a.key < b.key ? -1 : a.key > b.key ? 1 : 0));

    const [firstStaff, secondStaff] = splitDay === "A" ? ["Jasmine", "Cathy"] : ["Cathy", "Jasmine"];
    const forcedOwner = (exJ, exC) => (exJ && !exC ? "Cathy" : exC && !exJ ? "Jasmine" : null);

    let reportable = 0, forcedFirst = 0, free = 0;
    for (const p of pool) {
      if (p.exJ && p.exC) continue;
      reportable++;
      const owner = forcedOwner(p.exJ, p.exC);
      if (owner === null) free++;
      else if (owner === firstStaff) forcedFirst++;
    }
    const firstHalfFree = Math.min(Math.max(Math.floor(reportable / 2) - forcedFirst, 0), free);

    let freeSeen = 0;
    const poolStaff = pool.map((p) => {
      const inFirst = freeSeen < firstHalfFree;
      let owner = forcedOwner(p.exJ, p.exC);
      if (owner === null) {
        owner = inFirst ? firstStaff : secondStaff;
        if (!(p.exJ && p.exC)) freeSeen++;
      }
      return owner;
    });
    if (o.skipCathy) poolStaff.fill("Jasmine");

    pool.forEach((p, i) => { p.row[0] = cell(poolStaff[i], p.row[0] && p.row[0].s); });
    others.forEach((r) => { r[0] = cell(fixed.get(r), r[0] && r[0].s); });
    table.rows = pool.map((p) => p.row).concat(others);
    return splitDay;
  }

  /**
   * process_workbook (app.py): the whole run. Returns everything needed to
   * write the output workbooks. `table` is modified in place into the Masters.
   */
  function processTable(table, filename, opts) {
    opts = opts || {};
    const exclusions = new ReportExclusions(opts);
    if (opts.customReportName) validateCustomReportName(opts.customReportName);

    const dateToken = extractDateFromFilename(filename);
    if (!dateToken) throw new Error("Filename must contain 8 digits (MMDDYYYY)");
    const prefix = getFilenamePrefix(filename);

    const invalidRows = extractInvalid(table);
    const invalidHeaders = table.headers.map((h) => cell(h.v));
    const splitDay = assignStaff(table, dateToken, {
      includeProgramming: opts.includeProgramming, skipCathy: opts.skipCathy, splitDay: opts.splitDay,
      exclusions, phpToJasmine: opts.phpToJasmine, customReportName: opts.customReportName,
      customReportPayerTerms: opts.customReportPayerTerms,
      customReportProfessionalOnly: opts.customReportProfessionalOnly,
    });

    // Export columns; a later duplicate header wins, like app.py.
    let serviceCol = 0, programLevelCol = 0, payerCol = 0, divisionCol = 0;
    table.headers.forEach((h, i) => {
      const n = plain(h.v);
      if (n === "Service") serviceCol = i + 1;
      else if (n === "Program Level") programLevelCol = i + 1;
      else if (n === "Payer") payerCol = i + 1;
      else if (n === "GROUPFLD1") divisionCol = i + 1;
    });

    const wmCount = programLevelCol
      ? table.rows.filter((r) => isWmProgramLevel(val(r, programLevelCol))).length : 0;

    const staffReports = opts.skipCathy ? ["Jasmine", "CB"] : ["Cathy", "Jasmine", "CB"];
    if (opts.customReportName) staffReports.push(opts.customReportName);

    const reports = [];
    for (const staff of staffReports) {
      const rows = table.rows.filter((r) => {
        if (plain(val(r, 1)) !== staff) return false;
        if (exclusions.excludes(staff,
          payerCol ? text(val(r, payerCol)) : "",
          serviceCol ? text(val(r, serviceCol)) : "",
          divisionCol ? text(val(r, divisionCol)) : "")) return false;
        if (staff !== "Melissa" && programLevelCol && isWmProgramLevel(val(r, programLevelCol))) return false;
        return true;
      });
      if (!rows.length) continue;
      reports.push({ staff, filename: `${prefix}${staff}.xlsx`, statusColumns: staff !== "CB", rows });
    }

    return {
      dateToken, prefix, splitDay, wmCount, invalidCount: invalidRows.length,
      invalidRows, invalidHeaders, reports,
      mastersFilename: `${prefix}Masters.xlsx`,
      statusItems: STATUS_ITEMS.concat(["Batch Billings", "IOP"]),
    };
  }

  return {
    plain, pyStr, text, parseTerms, matchesAnyTerm, isEcare, isDrugScreen, isPhpService,
    isWithdrawalManagementService, isProgrammingService, isIopService, isDayTreatmentService,
    billsLikeIop, isProfessionalClaimType, isWmProgramLevel, isNonBillableServiceForWeekday,
    ReportExclusions, validateCustomReportName, extractDateFromFilename, getFilenamePrefix,
    parseWeekdayFromToken, splitDayForDateToken, extractInvalid, assignStaff, processTable,
    RESERVED_STAFF_NAMES,
  };
});
