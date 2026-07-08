// Verifies the site-review UX + FIRMette ZIP batch in web/index.html:
//   - site names surfaced as map labels (Leaflet tooltips)
//   - clicking a row zooms to that site and selects it
//   - Prev/Next steps site-by-site (with wrap-around)
//   - the authoritative source layers draw on the map around the selected
//     site, with an on-map legend citing the exact layers
//   - per-row "Source" links anchor into sources.html
//   - "Download FIRMettes (ZIP)" drives the FEMA GP flow per site and
//     produces a valid ZIP (validated with Python's zipfile, including CRCs)
//   - sources.html itself loads and documents the layers
//
// Same stubbing rationale as verify-pdf-report.mjs (see its header): the
// network is stubbed with real captured MDOT/ACUB fixtures (mi-geom.json /
// acub-geom.json are live frame-ENVELOPE captures — the 0.75 mi Kalamazoo
// frame with 88 segments across 6 classes — so the drawn overlay is the
// realistic street grid); the FEMA GP flow's URL shapes + CORS were
// confirmed live via curl on 2026-07-03 (submitJob -> jobs/{id} ->
// results/OutputFile -> PDF, all with access-control-allow-origin echoed).
// Classification queries use esriGeometryPoint; the review overlay's frame
// queries use esriGeometryEnvelope — the stubs dispatch on that.
//
//   cd build/web-tests && npm install && node verify-review-ui.mjs
import { chromium } from "playwright-core";
import { readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const fixture = name => readFileSync(join(here, "fixtures", name), "utf8");
const miMeta = fixture("mi-meta.json"), miGeom = fixture("mi-geom.json");
const acubMeta = fixture("acub-meta.json"), acubGeom = fixture("acub-geom.json");
const FAKE_PDF = Buffer.from("%PDF-1.4\n1 0 obj<<>>endobj\ntrailer<<>>\n%%EOF\n(fake firmette for zip test)");

const PAGE = "file://" + join(here, "..", "..", "web", "index.html");
const SOURCES = "file://" + join(here, "..", "..", "web", "sources.html");
const CHROMIUM_PATH = process.env.PLAYWRIGHT_CHROMIUM_PATH
  || "/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell";

const browser = await chromium.launch({ executablePath: CHROMIUM_PATH });
const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });

const errors = [];
page.on("pageerror", e => errors.push("pageerror: " + e.message));
page.on("console", m => {
  if (m.type() === "error" && !m.text().includes("Failed to load resource")) errors.push("console: " + m.text());
});

await page.route("**/*", async route => {
  const url = route.request().url();
  if (url.startsWith("file://")) return route.continue();
  if (url.includes("arcgisonline.com")) return route.abort();   // basemap tiles: irrelevant here
  const json = { contentType: "application/json" };

  // FEMA FIRMette GP flow (URL shapes confirmed live via curl 2026-07-03)
  if (url.includes("/PrintFIRMette/submitJob")) return route.fulfill({ ...json, body: JSON.stringify({ jobId: "jTEST", jobStatus: "esriJobSubmitted" }) });
  if (url.includes("/PrintFIRMette/jobs/jTEST/results/OutputFile")) return route.fulfill({ ...json, body: JSON.stringify({ paramName: "OutputFile", dataType: "GPDataFile", value: { url: "https://msc.fema.gov/fakeout/firmette-test.pdf" } }) });
  if (url.includes("/PrintFIRMette/jobs/jTEST")) return route.fulfill({ ...json, body: JSON.stringify({ jobStatus: "esriJobSucceeded" }) });
  if (url.includes("/fakeout/firmette-test.pdf")) return route.fulfill({ contentType: "application/pdf", body: FAKE_PDF });

  // classification pass (rr-core: point queries, WITH geometry since the
  // PR #24 port — distances are computed from each feature's polyline).
  // Site A (Kalamazoo, 42.28536,-85.57025): one Minor Collector under the
  // point + exact urban hit -> RED "Federal aid - Urban Minor Collector".
  // Site B (42.6911,-84.5360): a Local road ~2 ft away and a Major Collector
  // ~18 ft away -> the closest is non-federal but a federal road is within
  // 30 ft -> YELLOW "Review - Second road close".
  if (url.includes("FeatureServer/353/query") && url.includes("esriGeometryPoint")) {
    if (url.includes("-84.536"))
      return route.fulfill({ ...json, body: JSON.stringify({ features: [
        { attributes: { FunctionalSystem: 7, PR: "0343402" }, geometry: { paths: [[[-84.53601, 42.69105], [-84.53599, 42.69115]]] } },
        { attributes: { FunctionalSystem: 5, PR: "0343499" }, geometry: { paths: [[[-84.53607, 42.69105], [-84.53607, 42.69115]]] } },
      ] }) });
    return route.fulfill({ ...json, body: JSON.stringify({ features: [
      { attributes: { FunctionalSystem: 6, PR: "0006904" }, geometry: { paths: [[[-85.57026, 42.28530], [-85.57024, 42.28542]]] } },
    ] }) });
  }
  if (url.includes("FeatureServer/543/query")) return route.fulfill({ ...json, body: JSON.stringify({ features: [] }) });
  if (url.includes("NTAD_Adjusted_Urban_Areas/FeatureServer/0/query") && url.includes("esriGeometryPoint"))
    return route.fulfill({ ...json, body: JSON.stringify({ features: [{ attributes: { NAME: "Kalamazoo, MI", UACE: "43723", state_1: "MI" } }] }) });
  if (url.includes("TIGERweb")) return route.fulfill({ ...json, body: JSON.stringify({ features: [{ attributes: { NAME: "S Pitcher St" } }] }) });

  // review-overlay pass (frame envelope queries with geometry + layer metadata)
  if (url.includes("FeatureServer/353?f=json")) return route.fulfill({ ...json, body: miMeta });
  if (url.includes("FeatureServer/353/query") && url.includes("esriGeometryEnvelope")) return route.fulfill({ ...json, body: miGeom });
  if (url.includes("NTAD_Adjusted_Urban_Areas/FeatureServer/0?f=json")) return route.fulfill({ ...json, body: acubMeta });
  if (url.includes("NTAD_Adjusted_Urban_Areas/FeatureServer/0/query") && url.includes("esriGeometryEnvelope")) return route.fulfill({ ...json, body: acubGeom });

  return route.abort();
});

await page.goto(PAGE, { waitUntil: "domcontentloaded" });
await page.fill("#coordsIn", "Kalamazoo culvert,42.28536,-85.57025\nSite B,42.6911,-84.5360");
await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("2 point(s) classified"), { timeout: 15000 });

const checks = [];

// --- names on the map ---
const tooltipTexts = await page.locator(".leaflet-tooltip").allTextContents();
checks.push(["2 name labels on map", tooltipTexts.length === 2]);
checks.push(["labels carry the pasted names", tooltipTexts.join("|").includes("Kalamazoo culvert") && tooltipTexts.join("|").includes("Site B")]);

// --- result cards: verdict badge + per-segment class chips (PR #24 model) ---
const row0 = await page.locator("#resultsBody .row").first().textContent();
checks.push(["card carries FEDERAL AID badge text", row0.includes("FEDERAL AID")]);
checks.push(["card counts segments within the buffer, nearest first", row0.includes("1 road segment within 250 ft, nearest first")]);
checks.push(["card has a Minor Collector class chip", await page.locator("#resultsBody .row").first().locator(".chip", { hasText: "Minor Collector" }).count() === 1]);
checks.push(["closest chip is marked 'closest' with a distance", await page.evaluate(() => {
  const chip = document.querySelector("#resultsBody .row .chip.primary");
  return !!chip && chip.textContent.includes("closest") && /\(\d+ ft\)/.test(chip.textContent);
})]);
checks.push(["card cites the ACUB urban area", row0.includes("Urban · Kalamazoo, MI")]);
checks.push(["card lists TIGER street names in the merged road list", row0.includes("S Pitcher St")]);

// --- ambiguity path: Site B's closest road is Local but a Major Collector
// sits within 30 ft -> yellow "Review - Second road close" ---
const rowB = page.locator("#resultsBody .row").nth(1);
checks.push(["ambiguous site downgraded to REVIEW (yellow)", await rowB.evaluate(el => el.classList.contains("v-review"))]);
checks.push(["review reason 'Second road close' shown", (await rowB.textContent()).includes("Second road close")]);
checks.push(["review reason has a plain-language tooltip", await rowB.locator(".verdict-detail").getAttribute("title").then(t => (t || "").includes("30 ft"))]);
checks.push(["Site B shows both class chips with distances", await rowB.locator(".chip").count() >= 3
  && (await rowB.textContent()).includes("Major Collector")]);

// --- per-row map links: ArcGIS map (Excel AGOL NFC Layer parity) + Public map ---
checks.push(["primary ArcGIS map link uses MI curated webmap, pinned", await page.evaluate(() => {
  const a = [...document.querySelectorAll("#resultsBody .row a.maplink")];
  return a.length >= 2 && a.every(x => x.textContent === "ArcGIS map"
    && x.href.includes("webmap=6a1702b9147243d1a5ee62cd614bc681") && x.href.includes("marker="));
})]);
checks.push(["Public map link -> MI official app root (no coords)", await page.evaluate(() => {
  const a = [...document.querySelectorAll("#resultsBody .row a")].filter(x => x.textContent === "Public map");
  return a.length >= 2 && a.every(x => x.href.includes("experience.arcgis.com/experience/7edd160c205d46b481fcd605bb4c58ce")
    && !x.href.includes("marker"));
})]);

// --- search-radius (sensitivity) control ---
checks.push(["buffer select defaults to 250 ft", (await page.locator("#bufferSel").inputValue()) === "250"]);
await page.selectOption("#bufferSel", "50");
await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("2 point(s) classified"), { timeout: 15000 });
checks.push(["changing the radius re-classifies with the new buffer",
  (await page.locator("#resultsBody .row").first().textContent()).includes("within 50 ft")]);
await page.selectOption("#bufferSel", "250");
await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("2 point(s) classified"), { timeout: 15000 });

// --- GeoJSON export (drag-drop file for the ArcGIS / public maps) ---
{
  const [gj] = await Promise.all([
    page.waitForEvent("download", { timeout: 15000 }),
    page.click("#dlGeojson"),
  ]);
  checks.push(["geojson filename", gj.suggestedFilename() === "road-checker-sites.geojson"]);
  const gjData = JSON.parse(readFileSync(await gj.path(), "utf8"));
  const f0 = gjData.features && gjData.features[0], f1 = gjData.features && gjData.features[1];
  checks.push(["geojson: FeatureCollection with 2 point features", gjData.type === "FeatureCollection"
    && gjData.features.length === 2 && f0.geometry.type === "Point"
    && f0.geometry.coordinates[0] === -85.57025 && f0.geometry.coordinates[1] === 42.28536]);
  checks.push(["geojson: verdict + color + reason properties (Excel WriteSitesGeoJson parity)",
    f0.properties.Name === "Kalamazoo culvert" && f0.properties.Verdict === "Federal aid"
    && /^#/.test(f0.properties.VerdictColor) && f1.properties.ReviewNote === "Second road close"]);
}

// --- "How the colors are decided" explainer ---
checks.push(["how-colors explainer present", await page.evaluate(() => {
  const d = document.querySelector("details.howcolors");
  return !!d && d.textContent.includes("closest road decides red vs green") && d.textContent.includes("Yellow only downgrades green");
})]);

// --- click row -> zoom + select + layers ---
await page.locator("#resultsBody .row").nth(1).locator(".site-name").click();
await page.waitForFunction(() => document.getElementById("siteLegend").style.display === "block"
  && !document.getElementById("siteLegend").textContent.includes("loading"), { timeout: 15000 });
checks.push(["review info shows 2 / 2 + name", (await page.locator("#reviewInfo").textContent()).includes("2 / 2 — Site B")]);
checks.push(["clicked row is highlighted", await page.locator("#resultsBody .row.selected").count() === 1]);
checks.push(["map zoomed to site (z=17)", await page.evaluate(() => map.getZoom()) === 17]);
checks.push(["map centered on site", await page.evaluate(() => {
  const c = map.getCenter(); return Math.abs(c.lat - 42.6911) < 0.001 && Math.abs(c.lng - -84.5360) < 0.001; })]);

// --- source layers drawn + legend ---
checks.push(["source-layer geometry drawn on map (88-segment fixture)", await page.evaluate(() => siteOverlay.getLayers().length) > 50]);
const legendText = await page.locator("#siteLegend").textContent();
checks.push(["legend section header names the class layer", legendText.includes("Road functional class (MDOT)")]);
checks.push(["legend lists class labels from the source renderer", legendText.includes("Minor Collector") && legendText.includes("Local")]);
checks.push(["legend cites ACUB result", legendText.includes("Adjusted Urban Area: Kalamazoo, MI")]);
checks.push(["legend links to citations page", await page.locator('#siteLegend a[href^="sources.html#"]').count() === 1]);
// MI class + ACUB legend rows both link to MDOT's official Experience app
// MI legend shows the official MDOT app as the primary reference (canonical
// root URL, NO fragile hash deep-link) plus first-tier "pinned view" links
// into the FEMA Map Viewer that actually land on the site.
checks.push(["legend official reference → MDOT app canonical root (no hash)", await page.evaluate(() => {
  const a = [...document.querySelectorAll('#siteLegend a')].find(x => x.href.includes("7edd160c205d46b481fcd605bb4c58ce"));
  return !!a && !a.href.includes("widget_167") && !a.href.includes("#");
})]);
checks.push(["legend labels it as the Public map (official)", (await page.locator("#siteLegend").textContent()).includes("Public map (official)")]);
checks.push(["legend first-tier pinned links → FEMA Map Viewer", await page.locator('#siteLegend a[href*="fema.maps.arcgis.com"]').count() >= 2]);
checks.push(["legend pinned links labeled 'pinned view'", (await page.locator("#siteLegend").textContent()).includes("pinned view")]);

// --- next/prev stepping with wrap ---
await page.click("#nextSite");
await page.waitForFunction(() => document.getElementById("reviewInfo").textContent.includes("1 / 2"), { timeout: 10000 });
checks.push(["Next wraps to site 1", (await page.locator("#reviewInfo").textContent()).includes("Kalamazoo culvert")]);
await page.click("#prevSite");
checks.push(["Prev returns to site 2", (await page.locator("#reviewInfo").textContent()).includes("2 / 2")]);

// --- per-row Source links ---
checks.push(["rows link to sources.html#mi", await page.locator('#resultsBody a[href="sources.html#mi"]').count() >= 2]);

// --- FIRMette ZIP ---
const [download] = await Promise.all([
  page.waitForEvent("download", { timeout: 60000 }),
  page.click("#firmZipBtn"),
]);
checks.push(["zip filename", download.suggestedFilename() === "firmettes.zip"]);
const zipPath = await download.path();
const zipReport = execFileSync("python3", ["-c", `
import zipfile, sys, json
z = zipfile.ZipFile(sys.argv[1])
names = sorted(z.namelist())
bad = z.testzip()
starts = all(z.read(n).startswith(b"%PDF") for n in names)
print(json.dumps({"names": names, "bad": bad, "allPdf": starts}))
`, zipPath]).toString();
const zr = JSON.parse(zipReport);
checks.push(["zip contains 2 PDFs with site names", zr.names.length === 2
  && zr.names.includes("Kalamazoo culvert FIRMette.pdf") && zr.names.includes("Site B FIRMette.pdf")]);
checks.push(["zip CRCs valid (testzip clean)", zr.bad === null]);
checks.push(["zip entries are PDFs", zr.allPdf === true]);
checks.push(["firmette button restored", (await page.locator("#firmZipBtn").textContent()) === "Download FIRMettes (ZIP)"]);

// --- sources.html ---
await page.goto(SOURCES, { waitUntil: "domcontentloaded" });
const src = await page.content();
checks.push(["sources page: MI layer 353 documented", src.includes("NextGenPrFinderPub/FeatureServer/353")]);
checks.push(["sources page: IN record_status quirk", src.includes("record_status=5")]);
checks.push(["sources page: WI category-code quirk", src.includes("FNCT_CLS_CTGY_TYCD")]);
checks.push(["sources page: ACUB + FIRMette + TIGER sections", src.includes('id="acub"') && src.includes('id="firmette"') && src.includes('id="tiger"')]);
checks.push(["sources page: verdict-logic section (closest road, 30 ft rule, boundary edge)",
  src.includes('id="verdict"') && src.includes("Second road close") && src.includes("Urban boundary edge")
  && src.includes("closest road decides red vs green")]);

let fail = 0;
for (const [label, ok] of checks) { console.log((ok ? "  ok   " : "  FAIL ") + label); if (!ok) fail++; }
if (errors.length) { console.log("page errors:"); errors.forEach(e => console.log("  " + e)); fail++; }
console.log(fail ? "VERIFY FAILED" : "VERIFY PASSED");
await browser.close();
process.exit(fail ? 1 : 0);
