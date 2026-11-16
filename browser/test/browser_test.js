#!/usr/bin/env node
/*
 * Drive the built page in a real browser (Chromium via Playwright), the way a
 * person would: open dist/Billing App.html from disk, choose a file, tick
 * options, click Process, download every file.
 *   node test/browser_test.js INPUT.xlsx OUTDIR '{"phpToJasmine":true}' [screenshot.png]
 * Option keys are the element ids on the page (checkbox ids -> true, text
 * ids -> string, splitDay -> "A"/"B"). Prints a JSON summary, including any
 * network request the page tried to make (there should be none).
 */
const fs = require("fs");
const path = require("path");
const { chromium } = require("playwright");

(async () => {
  const [input, outdir, optsJson, shot] = process.argv.slice(2);
  const opts = optsJson ? JSON.parse(optsJson) : {};
  fs.mkdirSync(outdir, { recursive: true });
  const page_url = "file://" + path.resolve(__dirname, "..", "dist", "Billing App.html");
  const browser = await chromium.launch();
  const context = await browser.newContext({ acceptDownloads: true, viewport: { width: 900, height: 1400 } });
  const page = await context.newPage();
  const network = [], errors = [];
  page.on("request", (r) => { if (!/^(file|blob|data):/.test(r.url())) network.push(r.url()); });
  page.on("pageerror", (e) => errors.push(String(e)));
  page.on("console", (m) => { if (m.type() === "error") errors.push(m.text()); });

  await page.goto(page_url);
  await page.setInputFiles("#file", input);
  for (const [id, value] of Object.entries(opts)) {
    if (value === true) await page.check("#" + id);
    else if (value === false) await page.uncheck("#" + id);
    else if (id === "splitDay") await page.selectOption("#splitDay", value);
    else await page.fill("#" + id, String(value));
  }
  await page.click("#process");
  await page.waitForSelector("#results .downloads, #error .note", { timeout: 60000 });
  const error = await page.$eval("#error", (e) => e.innerText.trim());
  const files = [];
  if (!error) {
    const links = await page.$$("#results .downloads a");
    for (const a of links) {
      const [dl] = await Promise.all([page.waitForEvent("download"), a.click()]);
      const name = dl.suggestedFilename();
      await dl.saveAs(path.join(outdir, name));
      files.push(name);
    }
  }
  const text = await page.$eval("#results", (e) => e.innerText);
  if (shot) await page.screenshot({ path: shot, fullPage: true });
  await browser.close();
  console.log(JSON.stringify({ error, files, network, errors, resultText: text }, null, 1));
})();
