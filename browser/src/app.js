/* The page: reads the options, runs the processing, offers the downloads. */
(function () {
  "use strict";
  const Core = window.BillingCore;
  const IO = window.WorkbookIO;
  const MAX_FILE_SIZE = 50 * 1024 * 1024;
  const $ = (id) => document.getElementById(id);
  let currentFile = null;
  let downloadUrls = [];

  function options() {
    const split = $("splitDay").value;
    return {
      excludeOptum: $("excludeOptum").checked,
      excludeBcbAnthemCt: $("excludeBcbAnthemCt").checked,
      excludeAnthemCathyJasmine: $("excludeAnthem").checked,
      excludeDetoxResidential: $("excludeDetoxResidential").checked,
      includeProgramming: $("includeProgramming").checked,
      excludeAetna: $("excludeAetna").checked,
      skipCathy: $("skipCathy").checked,
      phpToJasmine: $("phpToJasmine").checked,
      splitDay: split || null,
      customReportName: $("customName").value.trim() || null,
      customReportPayerTerms: Core.parseTerms($("customPayers").value),
      customReportProfessionalOnly: $("customProfessionalOnly").checked,
      payerTerms: Core.parseTerms($("excludePayers").value),
      serviceTerms: Core.parseTerms($("excludeServices").value),
      scope: ["Jasmine", "Cathy", "CB"].filter((n) => $("scope" + n).checked),
      divisionPayerTerms: Core.parseTerms($("divisionPayers").value),
      divisionTerms: Core.parseTerms($("divisionDivisions").value),
    };
  }

  function updateWarnings() {
    const o = options();
    const w = [];
    if (o.includeProgramming && o.excludeDetoxResidential) {
      w.push("“Include Programming (Detox/Residential) today” and “Don't give anyone Detox or " +
        "Residential” are both checked. The exclusion wins: Programming rows will be assigned in the " +
        "Masters report but kept out of every individual workbook.");
    }
    if (o.customReportName && !o.customReportPayerTerms.length) {
      w.push("A custom report name is set but no payers were entered, so it won't claim any rows.");
    }
    if (o.customReportPayerTerms.length && !o.customReportName) {
      w.push("Custom report payers were entered but no staff name was given, so they'll be ignored.");
    }
    if (o.divisionPayerTerms.length && !o.divisionTerms.length) {
      w.push("A funding source is set to be removed but no division was entered, so this exclusion won't apply to anything.");
    }
    if (o.divisionTerms.length && !o.divisionPayerTerms.length) {
      w.push("A division is set but no funding source was entered, so this exclusion won't apply to anything.");
    }
    $("warnings").innerHTML = "";
    w.forEach((msg) => $("warnings").appendChild(note("warn", msg)));
  }

  function note(kind, html) {
    const d = document.createElement("div");
    d.className = "note " + kind;
    d.innerHTML = html;
    return d;
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  function setFile(file) {
    $("results").innerHTML = "";
    $("error").innerHTML = "";
    if (!file) { currentFile = null; $("fileName").textContent = "No file chosen"; $("process").disabled = true; return; }
    currentFile = file;
    $("fileName").textContent = file.name;
    $("process").disabled = false;
  }

  function clearDownloads() {
    downloadUrls.forEach((u) => URL.revokeObjectURL(u));
    downloadUrls = [];
  }

  async function processFile() {
    const file = currentFile;
    $("results").innerHTML = "";
    $("error").innerHTML = "";
    clearDownloads();
    if (!file) return;
    try {
      if (!/\.xlsx$/i.test(file.name)) throw new Error("Only .xlsx files are allowed");
      if (file.size > MAX_FILE_SIZE) throw new Error(`File too large (${(file.size / 1048576).toFixed(1)}MB). Maximum allowed: 50MB`);
      $("process").disabled = true;
      $("process").textContent = "Processing…";
      const data = await file.arrayBuffer();
      const { result, files } = await IO.run(window.ExcelJS, data, file.name, options());
      showResults(result, files);
      $("results").scrollIntoView({ behavior: "smooth", block: "start" });
    } catch (e) {
      $("error").appendChild(note("error", "<strong>Error:</strong> " + escapeHtml(e.message || e)));
      $("error").scrollIntoView({ behavior: "smooth", block: "center" });
    } finally {
      $("process").disabled = !currentFile;
      $("process").textContent = "Process file";
    }
  }

  function showResults(result, files) {
    const out = $("results");
    const o = options();
    if (result.wmCount > 0) {
      out.appendChild(note("warn", `Alert: ${result.wmCount} row(s) with 'WM' or 'OP WM' in the Program Level column ` +
        "were found. These rows are assigned to Melissa and appear only in the Masters report. Only Melissa is " +
        "authorized to bill WM."));
    }
    out.appendChild(note("ok", "<strong>✓ Processing complete!</strong>"));
    if (o.phpToJasmine) out.appendChild(note("info", "PHP: Insurance PHP rows went to <strong>Jasmine</strong> this run, not Melissa."));
    let split;
    if (o.skipCathy) split = `Split day: <strong>Day ${result.splitDay}</strong> — Cathy was skipped, so Jasmine has the whole shared pool.`;
    else if (result.splitDay === "A") split = "Split day: <strong>Day A</strong> — Jasmine has the A-M half, Cathy has the N-Z half.";
    else split = "Split day: <strong>Day B</strong> — Jasmine has the N-Z half, Cathy has the A-M half.";
    out.appendChild(note("info", split));

    const metrics = document.createElement("div");
    metrics.className = "metrics";
    [["Date Extracted", result.dateToken], ["Invalid Rows", result.invalidCount], ["Files Generated", files.length]]
      .forEach(([label, value]) => {
        const m = document.createElement("div");
        m.className = "metric";
        m.innerHTML = `<div class="label">${label}</div><div class="value">${escapeHtml(value)}</div>`;
        metrics.appendChild(m);
      });
    out.appendChild(metrics);

    const h = document.createElement("h2");
    h.textContent = "Download Results";
    out.appendChild(h);
    const list = document.createElement("div");
    list.className = "downloads";
    const links = files.map((f) => {
      const url = URL.createObjectURL(new Blob([f.buffer],
        { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      downloadUrls.push(url);
      const a = document.createElement("a");
      a.href = url;
      a.download = f.filename;
      a.className = "button secondary";
      a.textContent = "Download " + f.filename;
      list.appendChild(a);
      return a;
    });
    out.appendChild(list);
    const all = document.createElement("button");
    all.type = "button";
    all.className = "button";
    all.textContent = "Download all " + files.length + " files";
    all.addEventListener("click", async () => {
      for (const a of links) { a.click(); await new Promise((r) => setTimeout(r, 400)); }
    });
    out.appendChild(all);
    out.appendChild(note("muted", "Your browser may ask once whether this page can download several files — allow it. " +
      "Files go to your usual Downloads folder."));
  }

  document.addEventListener("DOMContentLoaded", () => {
    const input = $("file");
    input.addEventListener("change", () => setFile(input.files[0] || null));
    const drop = $("drop");
    ["dragenter", "dragover"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("over"); }));
    ["dragleave", "drop"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
    drop.addEventListener("drop", (e) => { if (e.dataTransfer.files[0]) setFile(e.dataTransfer.files[0]); });
    $("process").addEventListener("click", processFile);
    document.querySelectorAll("input, select").forEach((el) => {
      if (el.id !== "file") el.addEventListener("input", updateWarnings);
      if (el.id !== "file") el.addEventListener("change", updateWarnings);
    });
    updateWarnings();
    if (!window.ExcelJS) $("error").appendChild(note("error", "The Excel library didn't load, so this page can't process files."));
  });
})();
