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
  // find-on-map searches (TIGERweb boundaries + road-name search) — must
  // match BEFORE the generic TIGERweb street-name stub below
  if (url.includes("State_County/MapServer/0/query"))
    return route.fulfill({ ...json, body: JSON.stringify({ features: [{ attributes: { NAME: "Michigan" },
      geometry: { rings: [[[-90.4, 41.7], [-82.1, 41.7], [-82.1, 48.3], [-90.4, 48.3], [-90.4, 41.7]]] } }] }) });
  if (url.includes("State_County/MapServer/1/query")) {
    if (url.includes("returnGeometry=false"))
      return route.fulfill({ ...json, body: JSON.stringify({ features: [{ attributes: { NAME: "Kalamazoo County", GEOID: "26077", STATE: "26" } }] }) });
    return route.fulfill({ ...json, body: JSON.stringify({ features: [{ attributes: { NAME: "Kalamazoo County" },
      geometry: { rings: [[[-85.77, 42.16], [-85.42, 42.16], [-85.42, 42.42], [-85.77, 42.42], [-85.77, 42.16]]] } }] }) });
  }
  if (url.includes("Places_CouSub_ConCity_SubMCD/MapServer/1/query")) {
    if (url.includes("returnGeometry=false"))
      return route.fulfill({ ...json, body: JSON.stringify({ features: [{ attributes: { NAME: "Oshtemo charter township", GEOID: "2607761100", STATE: "26" } }] }) });
    return route.fulfill({ ...json, body: JSON.stringify({ features: [{ attributes: { NAME: "Oshtemo charter township" },
      geometry: { rings: [[[-85.70, 42.24], [-85.62, 42.24], [-85.62, 42.32], [-85.70, 42.32], [-85.70, 42.24]]] } }] }) });
  }
  if (url.includes("Transportation/MapServer/") && url.includes("UPPER(NAME)")) {
    if (url.includes("MapServer/8"))
      return route.fulfill({ ...json, body: JSON.stringify({ features: [
        { attributes: { NAME: "S Pitcher St" }, geometry: { paths: [[[-85.583, 42.284], [-85.583, 42.291]]] } },
        { attributes: { NAME: "N Pitcher St" }, geometry: { paths: [[[-85.583, 42.291], [-85.583, 42.298]]] } },
      ] }) });
    return route.fulfill({ ...json, body: JSON.stringify({ features: [] }) });
  }
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

// --- map pins take the row's verdict colors once classified ---
checks.push(["pins reflect the rows' verdict colors (red / yellow)", await page.evaluate(() =>
  JSON.stringify(currentPoints.filter(p => !p.invalid).map(p => p._marker.options.fillColor))
    === JSON.stringify([BUCKET_COLOR.fed, BUCKET_COLOR.review]))]);

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

// --- Auto-Detect header: disclaimer + Detection Buffer control ---
checks.push(["Auto-Detect heading, short disclaimer and Detection Buffer label with an info dot",
  await page.evaluate(() => {
    const h = [...document.querySelectorAll(".results h2")].map(x => x.textContent).join("|");
    const d = document.querySelector(".results .disclaim");
    const lbl = document.getElementById("bufferSel").closest("label").textContent;
    const info = document.querySelector(".results .infodot");
    return h.includes("Auto-Detect") && !h.includes("Results")
      && !!d && /verify/i.test(d.textContent)
      && lbl.includes("Detection Buffer") && !lbl.includes("Search buffer")
      && !!info && (info.title || "").length > 60;   // the buffer-logic explainer
  })]);

// --- detection-buffer (sensitivity) control ---
checks.push(["buffer select defaults to 250 ft", (await page.locator("#bufferSel").inputValue()) === "250"]);
await page.selectOption("#bufferSel", "50");
await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("2 point(s) classified"), { timeout: 15000 });
checks.push(["changing the radius re-classifies with the new buffer",
  (await page.locator("#resultsBody .row").first().textContent()).includes("within 50 ft")]);
await page.selectOption("#bufferSel", "250");
await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("2 point(s) classified"), { timeout: 15000 });

// --- Export pill + dialogue (every export lives in here now) ---
checks.push(["export pill sits with the input, not under Auto-Detect", await page.evaluate(() => {
  const pill = document.getElementById("exportBtn");
  const results = document.querySelector(".results");
  return !!pill && !results.contains(pill) && !!document.getElementById("exportCaret")
    && !document.querySelector(".sb-exports");
})]);
await page.click("#exportCaret");
checks.push(["dropdown lists both copy actions + one entry per format", await page.evaluate(() => {
  const t = [...document.querySelectorAll("#exportMenu button")].map(b => b.textContent).join("|");
  return !document.getElementById("exportMenu").hidden
    && t.includes("Copy site + coordinates") && t.includes("Copy Auto-Detect results")
    && t.includes("CSV file") && t.includes("KMZ") && t.includes("GeoJSON") && t.includes("PDF report");
})]);
await page.click('#exportMenu button[data-act="tab-geojson"]');
checks.push(["a menu format opens the dialogue on that tab, with a preview", await page.evaluate(() => {
  const open = !document.getElementById("exportWrap").hidden;
  const tab = document.querySelector('.extab[data-tab="geojson"]').classList.contains("active");
  const pane = !document.querySelector('.ex-pane[data-pane="geojson"]').hidden;
  return open && tab && pane && document.getElementById("prevGeojson").textContent.includes("FeatureCollection");
})]);
checks.push(["dialogue table is editable and carries a Note column", await page.evaluate(() => {
  const ths = [...document.querySelectorAll("#exportTable th")].map(t => t.textContent);
  const cells = document.querySelectorAll('#exportTable td[contenteditable="true"]');
  return ths[0] === "Site Name" && ths[ths.length - 1] === "Note" && cells.length >= ths.length;
})]);
// an edit in the table flows into the exports
await page.evaluate(() => {
  const tr = document.querySelector("#exportTable tr[data-k]");
  const td = tr.querySelectorAll("td")[EX.note];
  td.focus();
  td.textContent = "culvert undercut";
  td.dispatchEvent(new Event("input", { bubbles: true }));
});
const editedRowOk = await page.evaluate(() => exportRows()[0][EX.note] === "culvert undercut");
let editedPrevOk = false;   // the preview redraw is debounced ~250 ms
try {
  await page.waitForFunction(() =>
    document.getElementById("prevExcel").textContent.includes("culvert undercut")
    && document.getElementById("prevGeojson").textContent.includes("culvert undercut"), { timeout: 5000 });
  editedPrevOk = true;
} catch { /* reported as a failure below */ }
checks.push(["an edited cell flows into the export rows and the preview", editedRowOk && editedPrevOk]);
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
  await page.click("#exportClose");   // the overlay would block the row clicks below
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

// --- live layer mirror (browse mode): at z17 the viewport-wide MIRROR of the
// state class layer + ACUB should have fetched (same stubbed envelope
// fixtures), drawn into the canvas browse pane, and populated its own legend ---
await page.waitForFunction(() => {
  const el = document.getElementById("liveLegend");
  return el && el.style.display === "block" && el.textContent.includes("Urban boundary (USDOT NTAD 2020)")
    && !el.textContent.includes("Zoom to street level");
}, { timeout: 15000 });
const liveLegendText = await page.locator("#liveLegend").textContent();
checks.push(["live mirror legend lists source class labels", liveLegendText.includes("Minor Collector") && liveLegendText.includes("Local")]);
checks.push(["live mirror legend cites the urban-area layer", liveLegendText.includes("2020 Adjusted Urban Area")]);
checks.push(["live mirror drew into its canvas pane (under pins)", await page.evaluate(() =>
  !!document.querySelector(".leaflet-browse-pane canvas") && browseOverlay.getLayers().length > 50)]);
// progressive low-zoom band: at z11 the mirror fetches arterials-only
// (classCap 3 -> MI where clause carries FunctionalSystem <= 3) and the
// legend discloses the partial display
await page.evaluate(() => map.setView([42.28536, -85.57025], 11));
await page.waitForFunction(() => {
  const el = document.getElementById("liveLegend");
  return el && el.textContent.includes("principal arterials") && el.textContent.includes("zoom in for collectors");
}, { timeout: 15000 });
checks.push(["live mirror low-zoom band discloses arterials-only display", true]);
checks.push(["low-zoom band queried with a class-cap filter", await page.evaluate(() =>
  netLines.some(l => l.includes("FeatureServer/353/query") && l.includes("FunctionalSystem%20%3C%3D%203")))]);
await page.evaluate(() => map.setView([42.6911, -84.5360], 17));   // restore for the checks below
await page.waitForFunction(() => document.getElementById("liveLegend").textContent.includes("Minor Collector"), { timeout: 15000 });

checks.push(["unchecking Live layers clears the mirror", await page.evaluate(async () => {
  document.getElementById("liveLayers").checked = false;
  document.getElementById("liveLayers").dispatchEvent(new Event("change"));
  await new Promise(r => setTimeout(r, 50));
  const cleared = browseOverlay.getLayers().length === 0 && document.getElementById("liveLegend").style.display === "none";
  document.getElementById("liveLayers").checked = true;
  document.getElementById("liveLayers").dispatchEvent(new Event("change"));
  return cleared;
})]);

// The results table mirrors the map view now (the "in map view" checkbox is
// gone), and the legend checks above zoomed to one site — put every site back
// in view before the row-level checks below.
await page.evaluate(() => map.fitBounds(validPoints().map(p => [p.lat, p.lon]), { padding: [30, 30] }));

// --- full-page shell: sidebar pane + no page scrolling ---
checks.push(["full-page shell: sidebar pane, map fills the rest, no page scroll", await page.evaluate(() => {
  const sb = document.getElementById("sidebar"), ma = document.getElementById("mapArea");
  const noScroll = document.body.scrollHeight <= window.innerHeight + 1;
  return sb && ma && sb.offsetHeight >= window.innerHeight - 1
    && ma.getBoundingClientRect().right >= window.innerWidth - 1 && noScroll;
})]);

// --- compact rows: detail hidden until the row is expanded ---
checks.push(["rows are compact until clicked (detail + links hidden)", await page.evaluate(() => {
  const row = document.querySelector("#resultsBody .row:not(.open)") || document.querySelector("#resultsBody .row");
  const sub = row.querySelector(".row-sub"), links = row.querySelector(".row-links");
  const collapsedHidden = !row.classList.contains("open")
    ? (sub ? sub.offsetParent === null : true) && links.offsetParent === null : true;
  row.classList.add("open");
  const openShows = (sub ? sub.offsetParent !== null : true) && links.offsetParent !== null;
  row.classList.remove("open");
  return collapsedHidden && openShows;
})]);

// --- pop-out: the full-detail table over the map ---
checks.push(["pop-out table lists every site with full detail", await page.evaluate(() => {
  document.getElementById("popoutBtn").click();
  const wrap = document.getElementById("popoutWrap");
  const rows = document.querySelectorAll("#popoutTable tr");
  const txt = document.getElementById("popoutTable").textContent;
  const ok = !wrap.hidden && rows.length === 3   // header + 2 sites
    && txt.includes("Kalamazoo culvert") && txt.includes("Federal aid")
    && txt.includes("Urban area") && txt.includes("42.28536");
  document.getElementById("popoutClose").click();
  return ok && wrap.hidden;
})]);

// --- the two input tabs: find + collector live on the second tab ---
await page.click("#tabBtnCollect");
checks.push(["tabs: Search & Collect shows find + adder, hides the paste panel", await page.evaluate(() => {
  const collectShown = document.getElementById("collectPanel").offsetParent !== null
    && document.getElementById("findText").offsetParent !== null
    && document.getElementById("addPointBtn").offsetParent !== null;
  return collectShown && document.getElementById("inputPanel").offsetParent === null;
})]);

// --- find on map: type a state name -> county/township matches -> road search ---
await page.fill("#findText", "michigan");
await page.waitForFunction(() => [...document.querySelectorAll("#findResults .finditem")].some(d => d.textContent.includes("Michigan")), { timeout: 15000 });
await page.locator("#findResults .finditem", { hasText: "Michigan" }).first().click();
await page.waitForFunction(() => document.getElementById("reviewInfo").textContent.includes("Showing Michigan"), { timeout: 15000 });
checks.push(["find: typing a state name zooms to its boundary (no dropdown)", await page.evaluate(() =>
  !document.getElementById("findState") && findOverlay.getLayers().length > 0)]);
// type-ahead: filling the box (one input event) must surface suggestions
// after the debounce, with NO Find click
await page.fill("#findText", "kalamazoo");
// Wait for the REFRESHED results (the zoomed-out road note only the new
// term produces), not the previous term's still-displayed list.
await page.waitForFunction(() => {
  const t = document.getElementById("findResults").textContent;
  return t.includes("Kalamazoo County") && t.includes("Road-name search covers the visible map area");
}, { timeout: 15000 });
checks.push(["find: suggestions appear as you type (no Find click)", true]);
checks.push(["find: county + township matches listed with their kinds", await page.evaluate(() => {
  const t = [...document.querySelectorAll("#findResults .finditem")].map(d => d.textContent).join("|");
  return t.includes("Kalamazoo County") && t.includes("county") && t.includes("Oshtemo charter township") && t.includes("township");
})]);
{
  const frTxt = await page.locator("#findResults").textContent();
  const okZ = frTxt.includes("Road-name search covers the visible map area");
  if (!okZ) console.log("  findResults was:", frTxt.slice(0, 300), "| zoom:", await page.evaluate(() => map.getZoom()));
  checks.push(["find: statewide view explains road search needs zoom", okZ]);
}
await page.locator("#findResults .finditem", { hasText: "Kalamazoo County" }).click();
await page.waitForFunction(() => document.getElementById("reviewInfo").textContent.includes("Showing Kalamazoo County"), { timeout: 15000 });
checks.push(["find: county click zooms the map into the county", await page.evaluate(() => {
  const b = map.getBounds(); return b.getWest() > -86.5 && b.getEast() < -84.5 && map.getZoom() >= 9; })]);
checks.push(["'filter by map view' ships unchecked — the county view alone hides nothing",
  await page.evaluate(() => {
    const cb = document.getElementById("rowFilterView");
    return cb && !cb.checked && cb.closest("label").textContent.includes("filter by map view")
      && [...document.querySelectorAll("#resultsBody .row")].every(r => r.style.display !== "none");
  })]);
await page.check("#rowFilterView");
checks.push(["ticking it keeps only sites inside the county view", await page.evaluate(() =>
  [...document.querySelectorAll("#resultsBody .row")].filter(r => r.style.display !== "none").length === 1)]);
// A row click zooms to that site — which must NOT collapse the list to it.
await page.evaluate(() => map.fitBounds(validPoints().map(p => [p.lat, p.lon]), { padding: [30, 30] }));
await page.evaluate(() => selectSite(0));
checks.push(["zooming to one site under review keeps the other sites listed", await page.evaluate(() =>
  [...document.querySelectorAll("#resultsBody .row")].filter(r => r.style.display !== "none").length === 2)]);
await page.uncheck("#rowFilterView");
await page.fill("#findText", "pitcher");
await page.click("#findBtn");
await page.waitForFunction(() => [...document.querySelectorAll("#findResults .finditem")].some(d => d.textContent.includes("S Pitcher St")), { timeout: 15000 });
// road suggestions get annotated (async) with the state's FHWA class from
// the stubbed MDOT 353 point query (Minor Collector), swatched in class color
await page.waitForFunction(() => {
  const it = [...document.querySelectorAll("#findResults .finditem")].find(d => d.textContent.includes("S Pitcher St"));
  return it && it.textContent.includes("Minor Collector") && !!it.querySelector(".cw");
}, { timeout: 15000 });
checks.push(["find: road suggestion carries FHWA class + color swatch", true]);
await page.locator("#findResults .finditem", { hasText: "S Pitcher St" }).first().click();
await page.waitForFunction(() => document.getElementById("reviewInfo").textContent.includes("Showing road: S Pitcher St"), { timeout: 15000 });

// --- Search & Collect: add a named+noted GPS point, it classifies like any row ---
await page.fill("#addName", "Washout site");
await page.fill("#addCoords", "42.28536, -85.57025");
await page.fill("#addNote", "north shoulder undercut");
await page.click("#addPointBtn");
await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("3 point(s) classified"), { timeout: 15000 });
checks.push(["collector: added point classifies into the shared table", await page.evaluate(() => {
  const rows = [...document.querySelectorAll("#resultsBody .row")];
  const row = rows.find(r => r.textContent.includes("Washout site"));
  return rows.length === 3 && !!row && row.textContent.includes("north shoulder undercut")
    && !!row.querySelector("a.rmpt");
})]);
checks.push(["collector: note flows into CSV rows + pop-out table", await page.evaluate(() => {
  const csvRow = exportRows().find(r => r[0] === "Washout site");
  document.getElementById("popoutBtn").click();
  const po = document.getElementById("popoutTable").textContent;
  document.getElementById("popoutClose").click();
  return EXPORT_HEADERS[EXPORT_HEADERS.length - 1] === "Note"
    && csvRow && csvRow[csvRow.length - 1] === "north shoulder undercut"
    && po.includes("north shoulder undercut");
})]);
checks.push(["collector: KMZ machinery present (button + store-zip builder)", await page.evaluate(async () => {
  const zip = makeZip([{ name: "doc.kml", data: new TextEncoder().encode("<kml/>") }]);
  const head = new Uint8Array(await zip.slice(0, 2).arrayBuffer());
  return !!document.getElementById("dlKmz") && head[0] === 0x50 && head[1] === 0x4b;   // "PK"
})]);
await page.evaluate(() => { document.querySelector("#resultsBody a.rmpt").click(); });
await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("2 point(s) classified"), { timeout: 15000 });
checks.push(["collector: remove takes the point back out", await page.evaluate(() =>
  document.querySelectorAll("#resultsBody .row").length === 2)]);
await page.click("#tabBtnCoords");
checks.push(["find: road click highlights the matched segments", await page.evaluate(() => findOverlay.getLayers().length >= 1)]);

// --- results-row text filter ---
await page.evaluate(() => map.fitBounds(validPoints().map(p => [p.lat, p.lon]), { padding: [30, 30] }));
await page.fill("#rowFilter", "Site B");
checks.push(["row filter hides non-matching rows + counts", await page.evaluate(() =>
  [...document.querySelectorAll("#resultsBody .row")].filter(r => r.style.display !== "none").length === 1
  && document.getElementById("filterCount").textContent.includes("showing 1 of 2"))]);
checks.push(["row filter hides the non-matching site's map pin too (table and map mirror)",
  await page.evaluate(() => markerLayer.getLayers().length === 1)]);
await page.fill("#rowFilter", "");
checks.push(["clearing the filter restores all rows and pins", await page.evaluate(() =>
  [...document.querySelectorAll("#resultsBody .row")].every(r => r.style.display !== "none")
  && markerLayer.getLayers().length === 2)]);

// --- next/prev stepping with wrap ---
await page.click("#nextSite");
await page.waitForFunction(() => document.getElementById("reviewInfo").textContent.includes("1 / 2"), { timeout: 10000 });
checks.push(["Next wraps to site 1", (await page.locator("#reviewInfo").textContent()).includes("Kalamazoo culvert")]);
await page.click("#prevSite");
checks.push(["Prev returns to site 2", (await page.locator("#reviewInfo").textContent()).includes("2 / 2")]);

// --- Plot all sites: every site's frame drawn at once + combined legend ---
await page.click("#plotAllSites");
await page.waitForFunction(() => document.getElementById("reviewInfo").textContent.includes("All 2 site"), { timeout: 15000 });
await page.waitForFunction(() => !document.getElementById("siteLegend").textContent.includes("loading"), { timeout: 15000 });
checks.push(["plot-all deselects the single-site review", await page.locator("#resultsBody .row.selected").count() === 0]);
checks.push(["plot-all draws source layers around every site (2 x 88-segment fixture)",
  await page.evaluate(() => siteOverlay.getLayers().length) > 100]);
const allLegend = await page.locator("#siteLegend").textContent();
checks.push(["plot-all legend combines class labels + ACUB", allLegend.includes("all sites")
  && allLegend.includes("Minor Collector") && allLegend.includes("Adjusted Urban Area")]);

// --- per-row Source links ---
checks.push(["rows link to sources.html#mi", await page.locator('#resultsBody a[href="sources.html#mi"]').count() >= 2]);

// --- FIRMette ZIP (PDF tab of the export dialogue) ---
await page.evaluate(() => { document.getElementById("exportWrap").hidden = true; });
await page.click("#exportBtn");
await page.click('.extab[data-tab="pdf"]');
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
checks.push(["firmette button restored", (await page.locator("#firmZipBtn").textContent()) === "FIRMettes (ZIP)"]);
await page.click("#exportClose");

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
