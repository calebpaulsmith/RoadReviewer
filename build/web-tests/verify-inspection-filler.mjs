// Verifier for web/inspection/index.html (the Cat C Road/LWC single-file
// form filler). Opens the built page from file://, fills a representative
// field in every section, draws on the signature/initials/sketch pads,
// attaches a geotagged photo fixture (twice, so one photo lands in the
// page-4 grid and one on an appended page), generates the PDF, and then
// re-opens the output with the vendored pdf-lib to assert field values,
// checkbox states, page count, and that the EXIF GPS auto-fill ran.
//
// Run: npm install && node verify-inspection-filler.mjs
// (chromium path defaults to the CCR sandbox install; override with
//  CHROMIUM=/path/to/chromium)
import { chromium } from "playwright-core";
import { createRequire } from "module";
import { fileURLToPath } from "url";
import path from "path";
import fs from "fs";

const require = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));
const repo = path.resolve(here, "..", "..");
const page_url = "file://" + path.join(repo, "web", "inspection", "index.html");
const fixture = path.join(here, "fixtures", "site-photo-geotagged.jpg");
const outPdf = path.join(here, "inspection-filler-output.pdf");
const executablePath = process.env.CHROMIUM || "/opt/pw-browsers/chromium";

const TEXT = {
  "Applicant3": "City of Example",
  "PA ID3": "123-45678-00",
  "SI Date3": "09/03/2026",
  "SI Name3": "C. Smith",
  "Work Order3": "WO-1042",
  "Damage3": "DMG-77",
  "Year Built3": "1998",
  "Length3": "450 LF",
  "Notes3": "Two-lane gravel road over box culvert.",
  "Site9": "1",
  "Damage Components9": "Surface - gravel",
  "Location Addres11": "typo field check",
  "Quantity 11": "12",
  "% Complete 10": "50",
  "Mitigation Considerations3": "Water eroded the driving surface.",
  "Comment27": "Applicant will armor shoulders.",
  "Comment31": "No insurance on roads.",
  "Comment39": "None known.",
  "Additional Notes/Comments3": "Verified by automated test.",
};
const CHECK = ["Check Box147", "Check Box157", "Check Box161", "Check Box171",
               "Check Box178", "Check Box186", "Check Box190", "Check Box212"];

function fail(msg) { console.error("FAIL: " + msg); process.exitCode = 1; }
function ok(msg) { console.log("ok: " + msg); }

const browser = await chromium.launch({ executablePath });
const page = await browser.newPage();
page.on("pageerror", (e) => fail("page error: " + e));
await page.goto(page_url);

// open every section so inputs are interactable
await page.evaluate(() => document.querySelectorAll("details.sec").forEach(d => d.open = true));

// text fields + checkboxes drive the app's internal `fields` Map; reach them
// via label text is fragile, so set through the DOM elements the app made.
await page.evaluate(({ TEXT, CHECK }) => {
  // the app exposes no globals; find inputs by walking the same spec labels
  // is overkill — instead trigger through the Map by re-dispatching events.
  // fields Map isn't global, so use DOM: every input the app created is
  // inside #sections; match by its label text.
  const byLabel = {};
  document.querySelectorAll("#sections label.f").forEach(l => {
    const el = l.nextElementSibling;
    if (el) byLabel[l.firstChild.textContent.trim()] = el;
  });
  window.__byLabel = byLabel; // for debugging
}, { TEXT, CHECK });

// Simpler & robust: patch values via the app's own draft-import path.
const draft = { ...TEXT };
for (const c of CHECK) draft[c] = true;
await page.evaluate((draft) => {
  const blob = new Blob([JSON.stringify(draft)], { type: "application/json" });
  const file = new File([blob], "draft.json", { type: "application/json" });
  const dt = new DataTransfer();
  dt.items.add(file);
  const input = document.getElementById("draftFile");
  input.files = dt.files;
  input.dispatchEvent(new Event("change", { bubbles: true }));
}, draft);
await page.waitForTimeout(300);

// draw on the three pads (scroll each into the viewport first — mouse
// coordinates are viewport-relative, so an off-screen pad gets no ink)
for (const id of ["sigPad", "iniPad", "sketchPad"]) {
  const loc = page.locator("#" + id);
  await loc.scrollIntoViewIfNeeded();
  const box = await loc.boundingBox();
  await page.mouse.move(box.x + 10, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width - 10, box.y + box.height / 2, { steps: 8 });
  await page.mouse.move(box.x + box.width / 2, box.y + 10, { steps: 8 });
  await page.mouse.up();
  const inked = await page.evaluate((id) => {
    const c = document.getElementById(id);
    const d = c.getContext("2d").getImageData(0, 0, c.width, c.height).data;
    let n = 0;
    for (let i = 3; i < d.length; i += 4) if (d[i] > 0) n++;
    return n;
  }, id);
  if (inked < 100) fail(`pad ${id}: only ${inked} inked pixels after drawing`);
}
ok("drew signature, initials, sketch (ink verified on canvas)");

// attach the geotagged photo twice (grid photo + appended-page photo)
await page.setInputFiles("#photoFile", [fixture, fixture]);
await page.waitForTimeout(800);
const gpsLat = await page.evaluate(() =>
  [...document.querySelectorAll("#sections input")].map(i => i.value).find(v => v.startsWith("35.12")) || "");
if (gpsLat !== "35.123456") fail(`EXIF GPS auto-fill: expected 35.123456, got "${gpsLat}"`);
else ok("EXIF GPS auto-filled start latitude from photo");

// generate + capture download
const [download] = await Promise.all([
  page.waitForEvent("download", { timeout: 60000 }),
  page.click("#gen"),
]);
await download.saveAs(outPdf);
ok("download: " + download.suggestedFilename());
if (!/CatC_Road_LWC_DMG-77\.pdf/.test(download.suggestedFilename()))
  fail("unexpected filename " + download.suggestedFilename());

// ---- Survey123 CSV import: fresh page, import the 2-record fixture,
// pick record 1 from the chooser, assert the mapped draft state ----
const page2 = await browser.newPage();
page2.on("pageerror", (e) => fail("csv page error: " + e));
await page2.goto(page_url);
await page2.evaluate(() => { try { localStorage.clear(); } catch (e) {} });
await page2.reload();
await page2.setInputFiles("#s123File", path.join(here, "fixtures", "survey123-export.csv"));
await page2.waitForSelector("#s123Choose:not([hidden])", { timeout: 5000 });
await page2.selectOption("#s123Select", "0");
await page2.click("#s123Pick");
await page2.waitForTimeout(300);
const draftState = await page2.evaluate(() =>
  JSON.parse(localStorage.getItem("catc-road-lwc-draft-v1") || "{}"));
const expectCsv = {
  "Applicant3": "City of Example",
  "SI Date3": "09/03/2026",
  "Damage3": "DMG-77",
  "GPS Latitude Start3": "35.123456",
  "GPS Longitude Start3": "-92.654321",
  "Site9": "1",
  "Damage Components9": "Surface - gravel",
  "Mitigation Considerations3": "Water eroded the driving surface.",
  "Comment39": "None observed on site.",
};
for (const [k, v] of Object.entries(expectCsv))
  if (draftState[k] !== v) fail(`csv import "${k}": expected "${v}", got "${draftState[k]}"`);
for (const cb of ["Check Box147", "Check Box150", "Check Box151", "Check Box157",
                  "Check Box158", "Check Box160", "Check Box171", "Check Box176",
                  "Check Box212"])
  if (draftState[cb] !== true) fail(`csv import: ${cb} should be checked`);
for (const cb of ["Check Box148", "Check Box149", "Check Box152", "Check Box153",
                  "Check Box159", "Check Box161", "Check Box210", "Check Box211"])
  if (draftState[cb] === true) fail(`csv import: ${cb} should NOT be checked`);
ok("Survey123 CSV import mapped text, dates, choices, multi-selects, and x/y GPS");
await browser.close();

// ---- verify the produced PDF with the vendored pdf-lib ----
const { PDFDocument } = require(path.join(repo, "build", "inspection-filler", "vendor", "pdf-lib.min.js"));
const doc = await PDFDocument.load(fs.readFileSync(outPdf), { updateMetadata: false });
const form = doc.getForm();

if (doc.getPageCount() !== 8) fail(`page count: expected 8 (7 form + 1 photo page), got ${doc.getPageCount()}`);
else ok("page count 8 (extra photo page appended)");

for (const [name, want] of Object.entries(TEXT)) {
  const got = form.getTextField(name).getText() || "";
  if (got !== want) fail(`text "${name}": expected "${want}", got "${got}"`);
}
ok("all text fields round-tripped (incl. typo'd names)");

for (const name of CHECK) {
  if (!form.getCheckBox(name).isChecked()) fail(`checkbox "${name}" not checked`);
}
if (form.getCheckBox("Check Box148").isChecked()) fail("Check Box148 should be unchecked");
ok("checkboxes round-tripped");

if (form.getTextField("Page of1").getText() !== "8") fail("total page field should be 8");
if (form.getTextField("Page of14").getText() !== "3") fail("page-3 number field should be 3");
ok("page-number fields auto-filled");

// DMS encoding in the fixture rounds at the centisecond, so compare within
// ~1e-5 degrees (about a meter) rather than string-exact.
// image stamps: page 1 signature, page 2 sketch, page 4 photo, page 8 extra
// photo must each have left an image XObject in the page resources
const { PDFName, PDFDict } = require(path.join(repo, "build", "inspection-filler", "vendor", "pdf-lib.min.js"));
for (const [idx, what] of [[0, "signature"], [1, "sketch"], [3, "photo grid"], [7, "appended photo"]]) {
  const res = doc.getPage(idx).node.Resources();
  const xo = res && res.lookupMaybe(PDFName.of("XObject"), PDFDict);
  const n = xo ? xo.entries().length : 0;
  if (n < 1) fail(`page ${idx + 1}: no image XObject (${what} missing)`);
}
ok("signature, sketch, and photo images present on their pages");

const lat = parseFloat(form.getTextField("GPS Latitude Start3").getText());
const lon = parseFloat(form.getTextField("GPS Longitude Start3").getText());
if (Math.abs(lat - 35.123456) > 1e-5 || Math.abs(lon - -92.654321) > 1e-5)
  fail(`GPS in PDF: ${lat}, ${lon}`);
else ok("EXIF GPS coordinates landed in the PDF");

console.log(process.exitCode ? "\nVERIFY FAILED" : "\nVERIFY PASSED");
