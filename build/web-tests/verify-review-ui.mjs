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
// /opt/pw-browsers/chromium is a stable symlink to the installed build; a
// version-pinned path (or playwright's own bundled default) rots whenever the
// sandbox image bumps, and fails with a misleading "npx playwright install".
const CHROMIUM_PATH = process.env.PLAYWRIGHT_CHROMIUM_PATH || "/opt/pw-browsers/chromium";

// IsolateSandboxedIframes puts a sandbox="allow-scripts" frame in its own
// process, where Playwright's request interception does not reach it (its
// JSONP request went straight to the network with the flag on: 0 route hits,
// ERR_CERT_AUTHORITY_INVALID from the sandbox proxy). Harness-only; the page
// itself is unchanged.
const browser = await chromium.launch({ executablePath: CHROMIUM_PATH, args: ["--disable-features=IsolateSandboxedIframes"] });
const page = await browser.newPage({ viewport: { width: 1500, height: 1000 } });

const errors = [];
const geocoderReqs = [];   // every request to the Census geocoder: url + which frame made it
page.on("request", r => { if (r.url().includes("geocoding.geo.census.gov")) geocoderReqs.push({ url: r.url(), main: r.frame() === page.mainFrame() }); });
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
  // --- address / road-name legs ---
  // Census geocoder, JSONP: the page reaches it only from a sandboxed frame.
  // 5201 Portage Rd -> a TIGER interpolation just off Portage Rd; 5300 -> a
  // spot where the class changes 150 ft south; "Zzz Way" -> a location whose
  // street name matches no Census road; "Nowhere" -> no match at all.
  if (url.includes("geocoding.geo.census.gov")) {
    const cb = (/callback=(\w+)/.exec(url) || [])[1] || "cb";
    const addr = decodeURIComponent((/address=([^&]+)/.exec(url) || ["", ""])[1]).replace(/\+/g, " ");
    const hit = (y, street, suf) => ({ result: { input: {}, addressMatches: [{ matchedAddress: addr.toUpperCase() + ", 49002",
      coordinates: { x: -85.560066, y }, tigerLine: { side: "L", tigerLineId: "12230681" },
      addressComponents: { streetName: street, suffixType: suf, preDirection: "", suffixDirection: "", city: "PORTAGE", state: "MI", zip: "49002" } }] } });
    const body = addr.includes("Nowhere") ? { result: { addressMatches: [] } }
      : addr.includes("Zzz") ? hit(42.24177, "ZZZ", "WAY")
      : addr.includes("5300") ? hit(42.240056, "PORTAGE", "RD")
      : hit(42.24177, "PORTAGE", "RD");
    return route.fulfill({ contentType: "application/javascript", body: `/**/${cb}(${JSON.stringify(body)});` });
  }
  // TIGER roads near the geocoded point (the snap query asks for SUFDIRABRV):
  // Airview Blvd crosses 20 ft away, Portage Rd runs N-S at lon -85.5601.
  if (url.includes("Transportation/MapServer/") && url.includes("SUFDIRABRV") && url.includes("esriGeometryPoint")) {
    if (!url.includes("MapServer/8")) return route.fulfill({ ...json, body: JSON.stringify({ features: [] }) });
    return route.fulfill({ ...json, body: JSON.stringify({ features: [
      { attributes: { NAME: "Airview Blvd", BASENAME: "Airview", SUFTYPEABRV: "Blvd", MTFCC: "S1400" }, geometry: { paths: [[[-85.5604, 42.24176], [-85.5598, 42.24176]]] } },
      { attributes: { NAME: "Portage Rd", BASENAME: "Portage", SUFTYPEABRV: "Rd", MTFCC: "S1400" }, geometry: { paths: [[[-85.56010, 42.2380], [-85.56010, 42.2450]]] } },
    ] }) });
  }
  // Road-name line "Portage Rd, Portage MI": the place resolves as a county
  // subdivision (count 1, then its extent), then two Portage Rd edges in it.
  if (url.includes("Places_CouSub_ConCity_SubMCD/MapServer/4/query")) {
    if (url.includes("returnCountOnly=true")) return route.fulfill({ ...json, body: JSON.stringify({ count: url.includes("KALAMAZOO") ? 1 : 0 }) });
    if (url.includes("returnExtentOnly=true")) return route.fulfill({ ...json, body: JSON.stringify({ extent: { xmin: -85.65, ymin: 42.20, xmax: -85.53, ymax: 42.33 } }) });
    if (url.includes("esriGeometryPoint")) return route.fulfill({ ...json, body: JSON.stringify({ features: [{ attributes: { NAME: "Portage city" } }] }) });
    if (url.includes("returnGeometry=true")) return route.fulfill({ ...json, body: JSON.stringify({ features: [{ attributes: { NAME: "Kalamazoo city" },
      geometry: { rings: [[[-85.65, 42.20], [-85.53, 42.20], [-85.53, 42.33], [-85.65, 42.33], [-85.65, 42.20]]] } }] }) });
    return route.fulfill({ ...json, body: JSON.stringify({ features: url.includes("KALAMAZOO") ? [{ attributes: { NAME: "Kalamazoo city", GEOID: "2642160", STATE: "26" } }] : [] }) });
  }
  if (url.includes("returnCountOnly=true")) return route.fulfill({ ...json, body: JSON.stringify({ count: url.includes("Places_CouSub_ConCity_SubMCD/MapServer/1/") ? 1 : 0 }) });
  if (url.includes("returnExtentOnly=true")) return route.fulfill({ ...json, body: JSON.stringify({ extent: { xmin: -85.65, ymin: 42.15, xmax: -85.53, ymax: 42.25 } }) });
  if (url.includes("Transportation/MapServer/") && url.includes("SUFDIRABRV") && url.includes("esriGeometryEnvelope")) {
    if (!url.includes("MapServer/8")) return route.fulfill({ ...json, body: JSON.stringify({ features: [] }) });
    return route.fulfill({ ...json, body: JSON.stringify({ features: [
      { attributes: { NAME: "Portage Rd", BASENAME: "Portage", SUFTYPEABRV: "Rd" }, geometry: { paths: [[[-85.5601, 42.2300], [-85.5601, 42.2400]]] } },
      { attributes: { NAME: "Portage Rd", BASENAME: "Portage", SUFTYPEABRV: "Rd" }, geometry: { paths: [[[-85.5601, 42.2400], [-85.5601, 42.2450]]] } },
    ] }) });
  }
  // The state's class layer around Portage Rd: Local (7) south of 42.2400,
  // Minor Collector (6) north of it — one class change, placed so that the
  // 5201 address and its ±150 ft samples all read 6, the 5300 address reads
  // 6 at its snap but 7 at the sample 150 ft south, and the road's two edge
  // midpoints (42.235 / 42.2425) read 7 and 6.
  if (url.includes("FeatureServer/353/query") && url.includes("esriGeometryPoint") && url.includes("-85.5601")) {
    const m = /geometry=(-?[\d.]+),(-?[\d.]+)/.exec(url); const lon = +m[1], lat = +m[2];
    return route.fulfill({ ...json, body: JSON.stringify({ features: [{ attributes: { FunctionalSystem: lat < 42.2400 ? 7 : 6, PR: "0006904" },
      geometry: { paths: [[[lon, lat - 0.0004], [lon, lat + 0.0004]]] } }] }) });
  }
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

// --- Quick Export menu + View and Export (the pane expands over the map; no modal) ---
checks.push(["Quick Export button sits with the input, not under Auto-Detect; no split caret, no modal wrappers", await page.evaluate(() => {
  const b = document.getElementById("exportBtn");
  return !!b && /Quick Export/.test(b.textContent) && !document.querySelector(".results").contains(b)
    && !document.getElementById("exportCaret") && !document.getElementById("popoutWrap") && !document.getElementById("exportWrap");
})]);
await page.click("#exportBtn");
checks.push(["Quick Export lists both copy actions + a direct download per format", await page.evaluate(() => {
  const t = [...document.querySelectorAll("#exportMenu button")].map(b => b.textContent).join("|");
  return !document.getElementById("exportMenu").hidden && t.includes("Copy site + coordinates") && t.includes("Copy Auto-Detect results")
    && t.includes("Download CSV") && t.includes("KMZ") && t.includes("GeoJSON") && t.includes("PDF report") && t.includes("FIRMettes");
})]);
{
  const [dl] = await Promise.all([page.waitForEvent("download", { timeout: 15000 }), page.click('#exportMenu button[data-act="dl-geojson"]')]);
  checks.push(["Quick Export: a format entry downloads straight away, no dialogue", dl.suggestedFilename() === "road-checker-sites.geojson"
    && await page.evaluate(() => document.getElementById("exportMenu").hidden && !document.body.classList.contains("expanded"))]);
}
await page.click("#viewExportBtn");
await page.waitForFunction(() => document.body.classList.contains("expanded") && document.getElementById("mapArea").offsetWidth < window.innerWidth * 0.3, { timeout: 5000 });
checks.push(["View and Export expands the pane over the map (map keeps ~1/6, card + table shown)", await page.evaluate(() => {
  const m = document.getElementById("mapArea").offsetWidth, w = window.innerWidth;
  return m > 150 && m < w * 0.3 && document.getElementById("exportCard").offsetParent !== null
    && document.getElementById("exportTable").querySelectorAll("tr").length === 3;
})]);
await page.click('.extab[data-tab="geojson"]');
checks.push(["a format tab shows its preview inside the expanded pane", await page.evaluate(() => {
  const tab = document.querySelector('.extab[data-tab="geojson"]').classList.contains("active");
  const pane = !document.querySelector('.ex-pane[data-pane="geojson"]').hidden;
  return tab && pane && document.getElementById("prevGeojson").textContent.includes("FeatureCollection");
})]);
checks.push(["clicking a table row selects the site: blue outline, synced with the list row, map zoomed in", await page.evaluate(async () => {
  const tr = document.querySelectorAll("#exportTable tr[data-k]")[1];
  tr.querySelector("td").click();
  await new Promise(r => setTimeout(r, 80));
  const p = validPoints()[reviewIdx];
  return tr.classList.contains("selected") && !!p && editKey(p) === tr.dataset.k && map.getZoom() >= 16
    && document.querySelector("#resultsBody .row.selected") === p._tr;
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
// --- identity columns write BACK to the coordinates box and re-classify ---
{
  const before = await page.inputValue("#coordsIn");
  await page.evaluate(() => {
    const tr = document.querySelector("#exportTable tr[data-k]");
    const td = tr.querySelectorAll("td")[EX.lon];
    td.focus();
    td.textContent = "-84.5360";                 // move site 1 onto site 2's street
    td.dispatchEvent(new Event("input", { bubbles: true }));
    td.blur();
  });
  await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("2 point(s) classified"), { timeout: 20000 });
  const after = await page.inputValue("#coordsIn");
  checks.push(["editing Longitude rewrites that line in the coordinates box, keeping the rest",
    after !== before && after.split("\n")[0].includes("-84.5360") && after.split("\n")[0].includes("42.28536")
    && after.split("\n")[1] === before.split("\n")[1]]);
  checks.push(["the moved site re-classifies and its pin moves with it", await page.evaluate(() => {
    const p = validPoints()[0];
    return p.lon === -84.536 && !!p.result && Math.abs(p._marker.getLatLng().lng + 84.536) < 1e-6;
  })]);
  // a coordinate the parser would never accept is refused, not exported
  await page.evaluate(() => {
    const tr = document.querySelector("#exportTable tr[data-k]");
    const td = tr.querySelectorAll("td")[EX.lat];
    td.focus();
    td.textContent = "not a latitude";
    td.dispatchEvent(new Event("input", { bubbles: true }));
    td.blur();
  });
  checks.push(["a bad coordinate is refused and the cell reverts", await page.evaluate(() =>
    document.querySelector("#exportTable tr[data-k]").querySelectorAll("td")[EX.lat].textContent === "42.28536"
    && exportRows()[0][EX.lat] === 42.28536)]);
  // Both halves of a transposed pair: TAB from Latitude to Longitude must
  // stash the first edit without re-checking a half-corrected location, then
  // leaving the row applies both at once.
  const boxBeforePair = await page.inputValue("#coordsIn");
  await page.evaluate(() => {
    const tr = document.querySelector("#exportTable tr[data-k]");
    const tds = tr.querySelectorAll("td");
    const latTd = tds[EX.lat], lonTd = tds[EX.lon];
    latTd.focus();
    latTd.textContent = "42.6911";
    latTd.dispatchEvent(new Event("input", { bubbles: true }));
    lonTd.focus();                      // tab to the next cell in the SAME row
  });
  checks.push(["tabbing to the next cell in the row saves the edit but holds the re-check",
    await page.evaluate(([box]) => document.getElementById("coordsIn").value === box
      && validPoints()[0].lat !== 42.6911, [boxBeforePair])]);
  await page.evaluate(() => {
    const tr = document.querySelector("#exportTable tr[data-k]");
    const lonTd = tr.querySelectorAll("td")[EX.lon];
    lonTd.textContent = "-84.5360";
    lonTd.dispatchEvent(new Event("input", { bubbles: true }));
    lonTd.blur();                       // leaving the row applies both
  });
  let bothOk = false;
  try {
    await page.waitForFunction(() => {
      const p = validPoints()[0];
      return p.lat === 42.6911 && p.lon === -84.536 && !!p.result;
    }, { timeout: 20000 });
    bothOk = true;
  } catch { /* reported below */ }
  checks.push(["leaving the row applies both halves of the pair in one re-check", bothOk]);

  // put it back so the checks below see the original two sites
  await page.evaluate(() => {
    const tr = document.querySelector("#exportTable tr[data-k]");
    const set = (col, v) => {
      const td = tr.querySelectorAll("td")[col];
      td.focus(); td.textContent = v;
      td.dispatchEvent(new Event("input", { bubbles: true }));
      td.blur();
    };
    set(EX.lat, "42.28536");
    set(EX.lon, "-85.57025");
  });
  await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("2 point(s) classified")
    && validPoints()[0].lat === 42.28536 && validPoints()[0].lon === -85.57025, { timeout: 20000 });
}

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
  return !!d && d.textContent.includes("closest road decides red vs blue") && d.textContent.includes("Amber only downgrades blue");
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
// Source provenance (2026-09-17): live-mirror legend rows link to the state
// DOT's PUBLIC map app (never the REST layer) and name MDOT on hover.
checks.push(["live mirror legend rows link to MDOT's official public map with a source tooltip", await page.evaluate(() => {
  const rows = [...document.querySelectorAll("#liveLegend a.lrow")].filter(a => a.href === PUBLIC_MAP.MI || a.href.startsWith(PUBLIC_MAP.MI));
  return rows.length >= 2 && rows.every(a => /MDOT/.test(a.title) && a.target === "_blank" && !/\/rest\//.test(a.href));
})]);
checks.push(["every classified row shows a Source chip that links to a public site", await page.evaluate(() => {
  const rows = [...document.querySelectorAll("#resultsBody .row")].filter(r => r.querySelector(".chip"));
  return rows.length > 0 && rows.every(r => { const a = r.querySelector("a.srcchip"); return a && /^Source: /.test(a.textContent) && /^https:/.test(a.href) && !/\/rest\//.test(a.href); });
})]);
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

// --- the results-header button opens the same expanded View and Export; Back returns the map ---
checks.push(["results-header 'View and Export' expands the pane with every site; Back restores the map", await page.evaluate(async () => {
  document.getElementById("popoutBtn").click();
  await new Promise(r => setTimeout(r, 480));
  const txt = document.getElementById("exportTable").textContent;
  const open = document.body.classList.contains("expanded") && document.querySelectorAll("#exportTable tr").length === 3
    && txt.includes("Kalamazoo culvert") && txt.includes("Federal aid") && txt.includes("42.28536");
  document.getElementById("exportClose").click();
  await new Promise(r => setTimeout(r, 480));
  return open && !document.body.classList.contains("expanded") && document.getElementById("mapArea").offsetWidth > window.innerWidth * 0.4;
})]);

// --- the two input tabs: find + collector live on the second tab ---
checks.push(["one pane: the search box sits above the coordinates box — no tabs, no Find button, no collect form", await page.evaluate(() => {
  const find = document.getElementById("findText"), box = document.getElementById("coordsIn");
  return find && box && find.offsetParent !== null && box.offsetParent !== null
    && find.getBoundingClientRect().top < box.getBoundingClientRect().top
    && !document.getElementById("tabBtnCoords") && !document.getElementById("tabBtnCollect") && !document.getElementById("collectPanel")
    && !document.getElementById("findBtn") && !document.getElementById("addPointBtn");
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
checks.push(["find: results grouped under State / County / City / Township headings, the city included", await page.evaluate(() => {
  const hdrs = [...document.querySelectorAll("#findResults .findhdr")].map(d => d.textContent.replace(/ ·.*$/, ""));
  const t = [...document.querySelectorAll("#findResults .finditem")].map(d => d.textContent).join("|");
  return hdrs.indexOf("County") < hdrs.indexOf("City") && hdrs.indexOf("City") < hdrs.indexOf("Township")
    && t.includes("Kalamazoo County") && t.includes("Kalamazoo city") && t.includes("Oshtemo charter township");
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
checks.push(["find: the county becomes the default area (chip with ✕, remembered in this browser)", await page.evaluate(() => {
  const bar = document.getElementById("areaBar");
  return !bar.hidden && bar.textContent.includes("Kalamazoo County, MI") && !!document.getElementById("areaClear")
    && defaultArea && defaultArea.kind === "county" && defaultArea.extent && (localStorage.getItem("rr_default_area") || "").includes("Kalamazoo");
})]);
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
await page.fill("#findText", "pitcher");   // type-ahead: no Find button
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

// --- the pin button: drop a site on the map, name it at the pin, it classifies like any row ---
await page.click("#pinBtn");
checks.push(["pin: one click arms single-pin mode (button lit, pin cursor)", await page.evaluate(() =>
  pinMode === 1 && document.getElementById("pinBtn").classList.contains("armed") && document.getElementById("map").classList.contains("pin-mode"))]);
await page.evaluate(() => map.fire("click", { latlng: L.latLng(42.29, -85.56), originalEvent: {} }));
checks.push(["pin: the map click appends 'Point N, lat, lon' to the coordinates box and opens the name editor with the default selected", await page.evaluate(() => {
  const last = document.getElementById("coordsIn").value.split("\n").pop();
  const ed = document.querySelector(".site-label.editing .pinname");
  const sel = window.getSelection();
  return last === "Point 3, 42.29000, -85.56000" && !!ed && document.activeElement === ed && sel && sel.toString() === "Point 3" && pinMode === 0;
})]);
await page.keyboard.type("Washout site");
await page.keyboard.press("Enter");
await page.waitForFunction(() => !pinEditor && !document.querySelector(".site-label.editing"), { timeout: 5000 }).catch(() => {});
{
  const st = await page.evaluate(() => ({ last: document.getElementById("coordsIn").value.split("\n").pop(), editing: !!document.querySelector(".site-label.editing"), pinMode }));
  const ok = st.last === "Washout site, 42.29000, -85.56000" && !st.editing;
  if (!ok) console.log("  pin state was:", JSON.stringify(st));
  checks.push(["pin: typing replaces the default name in the box; Enter finishes the editor", ok]);
}
await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("3 point(s) classified"), { timeout: 15000 });
checks.push(["pin: the dropped site classifies into the shared table and carries its label on the map", await page.evaluate(() => {
  const row = [...document.querySelectorAll("#resultsBody .row")].find(r => r.textContent.includes("Washout site"));
  return !!row && [...document.querySelectorAll(".site-label")].some(t => t.textContent === "Washout site");
})]);
// right-click the new pin: Move (drag) rewrites its line, Delete removes line + pin
await page.evaluate(() => { const p = currentPoints.find(q => q.name === "Washout site"); p._marker.fire("contextmenu", { latlng: L.latLng(p.lat, p.lon), originalEvent: new MouseEvent("contextmenu") }); });
checks.push(["pin menu: right-click opens Move / Delete next to the pin", await page.evaluate(() =>
  !!document.querySelector(".pinmenu button[data-act=move]") && !!document.querySelector(".pinmenu button[data-act=delete]"))]);
await page.evaluate(() => document.querySelector(".pinmenu button[data-act=move]").click());
await page.evaluate(() => { pinMove.marker.setLatLng([42.2955, -85.5605]); pinMove.marker.fire("dragend"); });
checks.push(["pin menu: Move rewrites that line's coordinates in the box (name kept)", await page.evaluate(() =>
  document.getElementById("coordsIn").value.split("\n").pop() === "Washout site, 42.29550, -85.56050" && !pinMove)]);
await page.waitForFunction(() => { const p = currentPoints.find(q => q.name === "Washout site"); return !!p && p.lat === 42.2955 && !!p.result; }, { timeout: 15000 });
await page.evaluate(() => { const p = currentPoints.find(q => q.name === "Washout site"); p._marker.fire("contextmenu", { latlng: L.latLng(p.lat, p.lon), originalEvent: new MouseEvent("contextmenu") }); });
await page.evaluate(() => document.querySelector(".pinmenu button[data-act=delete]").click());
// (Leaflet fades a removed tooltip for ~200 ms before dropping its element)
await page.waitForFunction(() => ![...document.querySelectorAll(".site-label")].some(t => t.textContent === "Washout site"), { timeout: 5000 }).catch(() => {});
checks.push(["pin menu: Delete removes the line from the box and the pin from the map", await page.evaluate(() =>
  document.getElementById("coordsIn").value.split("\n").length === 2 && !currentPoints.some(q => q.name === "Washout site")
  && ![...document.querySelectorAll(".site-label")].some(t => t.textContent === "Washout site"))]);
await page.click("#pinBtn"); await page.click("#pinBtn");
await page.evaluate(() => map.fire("click", { latlng: L.latLng(42.30, -85.55), originalEvent: {} }));
await page.keyboard.press("Enter");
await page.evaluate(() => map.fire("click", { latlng: L.latLng(42.31, -85.54), originalEvent: {} }));
await page.keyboard.press("Enter");
checks.push(["pin: two clicks arm keep-dropping mode — several pins, mode stays on; Escape turns it off", await page.evaluate(() => {
  const lines = document.getElementById("coordsIn").value.split("\n");
  const on = pinMode === 2 && document.getElementById("pinBtn").classList.contains("multi") && lines.length === 4;
  document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true }));
  return on && pinMode === 0 && !document.getElementById("map").classList.contains("pin-mode");
})]);
checks.push(["collector: KMZ machinery present (button + store-zip builder)", await page.evaluate(async () => {
  const zip = makeZip([{ name: "doc.kml", data: new TextEncoder().encode("<kml/>") }]);
  const head = new Uint8Array(await zip.slice(0, 2).arrayBuffer());
  return !!document.getElementById("dlKmz") && head[0] === 0x50 && head[1] === 0x4b;   // "PK"
})]);
await page.fill("#coordsIn", "Kalamazoo culvert,42.28536,-85.57025\nSite B,42.6911,-84.5360");
await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("2 point(s) classified"), { timeout: 15000 });
checks.push(["collector: remove takes the point back out", await page.evaluate(() =>
  document.querySelectorAll("#resultsBody .row").length === 2)]);
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

// --- verdict bar chart: live counts, bars are filter toggles, Next follows the filter ---
checks.push(["verdict chart: three labelled bars with live counts (1 federal aid, 1 needs review, 0 non-federal)", await page.evaluate(() => {
  const el = document.getElementById("verdictChart"), bs = [...el.querySelectorAll("button[data-bucket]")];
  const val = b => +bs.find(x => x.dataset.bucket === b).querySelector(".val").textContent;
  const lbl = b => bs.find(x => x.dataset.bucket === b).querySelector(".lbl").textContent;
  return !el.hidden && bs.length === 3 && val("fed") === 1 && val("review") === 1 && val("nonfed") === 0
    && lbl("fed") === "Federal aid" && lbl("review") === "Needs review" && lbl("nonfed") === "Non-federal aid"
    && bs.map(b => b.dataset.bucket).join() === "fed,review,nonfed";
})]);
await page.click('#verdictChart button[data-bucket="fed"]');
checks.push(["clicking the Federal aid bar keeps only that site (row + pin) and marks the bar on", await page.evaluate(() =>
  [...document.querySelectorAll("#resultsBody .row")].filter(r => r.style.display !== "none").length === 1
  && document.querySelector("#resultsBody .row:not([style*='none'])").textContent.includes("Kalamazoo culvert")
  && markerLayer.getLayers().length === 1 && document.querySelector('#verdictChart button[data-bucket="fed"]').classList.contains("on")
  && document.getElementById("filterCount").textContent.includes("showing 1 of 2"))]);
await page.click("#nextSite"); await page.click("#nextSite");
checks.push(["Next steps only through the filtered sites", (await page.locator("#reviewInfo").textContent()).includes("Kalamazoo culvert")]);
await page.click('#verdictChart button[data-bucket="fed"]');
checks.push(["clicking the bar again clears the filter", await page.evaluate(() =>
  [...document.querySelectorAll("#resultsBody .row")].every(r => r.style.display !== "none") && markerLayer.getLayers().length === 2
  && !document.querySelector("#verdictChart button.on"))]);
await page.evaluate(() => selectSite(1));

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
await page.evaluate(() => closeExport());
await page.click("#viewExportBtn");
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

// --- addresses and road names in the same box ---
// Coordinates classify as always; a road line resolves by itself through
// TIGERweb; addresses are highlighted and WAIT for the one Geocode button.
const ADDR_TEXT = "Kalamazoo culvert,42.28536,-85.57025\n5201 Portage Rd, Portage MI 49002\n5300 Portage Rd, Portage MI 49002\n" +
  "7 Zzz Way, Portage MI\n1 Nowhere Ln, Nowhere MI\nPortage Rd, Portage MI\ngarbage line";
geocoderReqs.length = 0;
await page.fill("#coordsIn", ADDR_TEXT);
await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("2 point(s) classified"), { timeout: 20000 });
await page.waitForFunction(() => [...document.querySelectorAll("#coordsMirror mark")].some(m => m.className.includes("road geo-ok")), { timeout: 20000 });
checks.push(["parse count names the kinds: 1 point, 4 addresses, 1 road, 1 unreadable",
  (await page.locator("#parseCount").textContent()) === "1 point(s) parsed, 4 addresses, 1 road, 1 line(s) unreadable"]);
checks.push(["address lines are highlighted in the box (4 marks), the road line too", await page.evaluate(() => {
  const marks = [...document.querySelectorAll("#coordsMirror mark")];
  return marks.filter(m => m.classList.contains("addr")).length === 4 && marks.filter(m => m.classList.contains("road")).length === 1
    && marks.find(m => m.classList.contains("addr")).textContent === "5201 Portage Rd, Portage MI 49002";
})]);
checks.push(["highlights sit under the typed lines (mirror and box share metrics)", await page.evaluate(() => {
  const box = document.getElementById("coordsIn"), mir = document.getElementById("coordsMirror");
  const cb = getComputedStyle(box), cm = getComputedStyle(mir);
  const rb = box.getBoundingClientRect(), rm = mir.getBoundingClientRect();
  return cb.fontFamily === cm.fontFamily && cb.fontSize === cm.fontSize && cb.lineHeight === cm.lineHeight && cb.paddingLeft === cm.paddingLeft
    && Math.abs(rb.left - rm.left) < 1 && Math.abs(rb.top - rm.top) < 1 && Math.abs(rb.width - rm.width) < 1;
})]);
checks.push(["one 'Geocode 4 addresses' button, shown only because addresses are present",
  !(await page.locator("#geoBar").isHidden()) && (await page.locator("#geocodeBtn").textContent()) === "Geocode 4 addresses"]);
checks.push(["NOTHING was sent to the geocoder on paste/parse", geocoderReqs.length === 0]);
checks.push(["the coordinate line classified without the button", (await page.locator("#resultsBody .row").first().textContent()).includes("FEDERAL AID")]);
const roadRow = page.locator("#resultsBody .row", { hasText: "Portage Rd (Portage)" });
checks.push(["the road-name line resolved by itself and classified the WHOLE road (mixed classes -> Review)", await roadRow.count() === 1
  && (await roadRow.textContent()).includes("whole road") && (await roadRow.textContent()).includes("Mixed classes on road")
  && (await roadRow.textContent()).includes("Local") && (await roadRow.textContent()).includes("Minor Collector")]);
checks.push(["address rows say they are waiting for the button", await page.locator("#resultsBody .row.v-address").count() === 4
  && (await page.locator("#resultsBody .row.v-address").first().textContent()).includes("not sent anywhere yet")]);

await page.click("#geocodeBtn");
await page.waitForFunction(() => !document.getElementById("geocodeBtn").disabled && document.getElementById("geocodeBtn").textContent.startsWith("Geocode"), { timeout: 40000 });
await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("5 point(s) classified"), { timeout: 20000 });
checks.push(["one click sent exactly 4 geocoder requests, all from a child (sandboxed) frame, all JSONP on the fixed host",
  geocoderReqs.length === 4 && geocoderReqs.every(r => !r.main && r.url.startsWith("https://geocoding.geo.census.gov/geocoder/locations/onelineaddress?")
    && r.url.includes("format=jsonp") && /callback=cb_g[a-z0-9]+/.test(r.url))]);
checks.push(["each request carries a distinct callback id", new Set(geocoderReqs.map(r => /callback=(\w+)/.exec(r.url)[1])).size === 4]);
checks.push(["every sandboxed frame is destroyed afterwards", await page.locator("iframe").count() === 0]);
checks.push(["the typed text is untouched — coordinates were never rewritten", (await page.inputValue("#coordsIn")) === ADDR_TEXT]);
const addr1 = page.locator("#resultsBody .row", { hasText: "5201 PORTAGE RD" });
checks.push(["5201 Portage Rd: snapped onto Portage Rd (not Airview Blvd 20 ft away) and classified from the snapped point",
  await addr1.count() === 1 && (await addr1.textContent()).includes("snapped to Portage Rd") && (await addr1.textContent()).includes("FEDERAL AID")
  && await page.evaluate(() => { const p = currentPoints.find(q => q.geo && q.name.startsWith("5201")); return !!p && p.lon === -85.5601 && p.geo.snapped && p.geo.distFt < 20 && p.result.verdict === "Federal aid - Urban Minor Collector"; })]);
const addr2 = page.locator("#resultsBody .row", { hasText: "5300 PORTAGE RD" });
checks.push(["5300 Portage Rd: class changes 150 ft along the street -> Review 'Class change nearby'",
  (await addr2.textContent()).includes("Class change nearby") && await addr2.evaluate(el => el.classList.contains("v-review"))]);
const addr3 = page.locator("#resultsBody .row", { hasText: "7 ZZZ WAY" });
checks.push(["a geocode whose street matches no Census road -> Review 'Street not matched', raw point kept",
  (await addr3.textContent()).includes("Street not matched") && (await addr3.textContent()).includes("street not matched")]);
checks.push(["no-match address stays an unresolved row with the reason, and the button offers it again",
  (await page.locator("#resultsBody .row", { hasText: "Nowhere" }).textContent()).includes("No match from the Census geocoder")
  && (await page.locator("#geocodeBtn").textContent()) === "Geocode 1 address"]);
checks.push(["mirror marks now show resolved (green) vs failed (red)", await page.evaluate(() => {
  const c = [...document.querySelectorAll("#coordsMirror mark")].map(m => m.className);
  return c.filter(x => x === "addr geo-ok").length === 3 && c.filter(x => x === "addr geo-bad").length === 1;
})]);
checks.push(["exports carry the snapped coordinates under the matched address", await page.evaluate(() => {
  const r = exportRows().find(x => String(x[0]).startsWith("5201 PORTAGE RD"));
  return !!r && r[1] === 42.24177 && r[2] === -85.5601;
})]);
// --- a bare road name resolves inside the default area; without one it says what to pick ---
await page.fill("#coordsIn", "Portage Rd");
await page.waitForFunction(() => [...document.querySelectorAll("#resultsBody .row")].some(r => r.textContent.includes("Kalamazoo County, MI")), { timeout: 20000 });
checks.push(["bare road name: looked up inside the default area and classified as a whole road", await page.evaluate(() => {
  const r = [...document.querySelectorAll("#resultsBody .row")].find(x => x.textContent.includes("Portage Rd (Kalamazoo County, MI)"));
  return !!r && r.textContent.includes("whole road");
})]);
await page.click("#areaClear");
await page.fill("#coordsIn", "Q Ave");   // a resolved line keeps its result; a NEW bare name has no area to look in
await page.waitForFunction(() => [...document.querySelectorAll("#resultsBody .row")].some(r => r.textContent.includes("Road name without a place")), { timeout: 10000 });
checks.push(["bare road name with no area: the row asks for a county / city / township, nothing is searched", await page.evaluate(() =>
  document.getElementById("areaBar").hidden && !localStorage.getItem("rr_default_area"))]);
await page.fill("#coordsIn", ADDR_TEXT);
await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("5 point(s) classified"), { timeout: 20000 });
checks.push(["clearing the addresses hides the button", await (async () => {
  await page.fill("#coordsIn", "Kalamazoo culvert,42.28536,-85.57025");
  return await page.waitForFunction(() => document.getElementById("geoBar").hidden, { timeout: 5000 }).then(() => true).catch(() => false);
})()]);
await page.fill("#coordsIn", "Kalamazoo culvert,42.28536,-85.57025\nSite B,42.6911,-84.5360");
await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("2 point(s) classified"), { timeout: 15000 });

// --- source outage (2026-09-17). This harness opens the page from file://,
// where the tilesets are never usable, so there is nothing to fall back to:
// the row fails, but the host is marked DOWN (strip + fast-fail) and, once
// the probe gets JSON again, the strip says it is back and offers to re-run
// the failed row. The tiles FALLBACK itself is covered by verify-hpms-tiles.
// (A route registered later is matched first, so this overrides the fixture.)
await page.route("**/mdotgis.state.mi.us/**", route => route.fulfill({ status: 500, contentType: "text/html", body: "<html>500</html>" }));
await page.fill("#coordsIn", "Outage check,42.28540,-85.57030");
await page.waitForFunction(() => /MDOT live layer/.test(document.getElementById("sourceStatus").textContent), { timeout: 20000 });
checks.push(["outage: status strip reports MDOT down with the reason and a recheck", await page.evaluate(() => {
  const s = document.getElementById("sourceStatus");
  const line = [...s.querySelectorAll("div")].map(d => d.textContent).find(t => /MDOT live layer/.test(t)) || "";
  return /HTTP 500/.test(line) && /rechecked/.test(line);
})]);
checks.push(["outage: a down host is not queried again (fast-fail, no request)", await page.evaluate(async () => {
  const before = netLines.filter(l => /GET .*mdotgis/.test(l)).length;
  try { await httpGetJson(svc("MI_NFC") + "/query?f=json"); return false; } catch (e) { if (!e.sourceDown) return false; }
  return netLines.filter(l => /GET .*mdotgis/.test(l)).length === before;
})]);
await page.waitForFunction(() => /1 point\(s\) classified/.test(document.getElementById("statusCount").textContent || ""), { timeout: 20000 });
checks.push(["outage (no tiles available): the row fails visibly with a retry link", await page.evaluate(() => {
  const r = document.querySelector("#resultsBody .row");
  return !!r && r.className.includes("v-failed") && !!r.querySelector("a.retry");
})]);
await page.route("**/mdotgis.state.mi.us/**", route => route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ name: "Functional System" }) }));
await page.evaluate(() => probeSource([...sourceHealth.keys()].find(h => /mdotgis/.test(h))));
await page.waitForFunction(() => /back up/.test(document.getElementById("sourceStatus").textContent), { timeout: 10000 });
checks.push(["outage: recovered source reported back up with a re-run offer for the failed row", await page.evaluate(() =>
  [...document.querySelectorAll("#sourceStatus div")].some(d => /MDOT live layer/.test(d.textContent) && /back up/.test(d.textContent) && /re-run 1 row/.test(d.textContent)))]);

// --- sources.html ---
await page.goto(SOURCES, { waitUntil: "domcontentloaded" });
const src = await page.content();
checks.push(["sources page: MI layer 353 documented", src.includes("NextGenPrFinderPub/FeatureServer/353")]);
checks.push(["sources page: IN record_status quirk", src.includes("record_status=5")]);
checks.push(["sources page: WI category-code quirk", src.includes("FNCT_CLS_CTGY_TYCD")]);
checks.push(["sources page: ACUB + FIRMette + TIGER sections", src.includes('id="acub"') && src.includes('id="firmette"') && src.includes('id="tiger"')]);
checks.push(["sources page: geocoder section says explicit-click only + interpolation caveat",
  src.includes('id="geocoder"') && src.includes("only when the") && src.includes("interpolation")]);
checks.push(["sources page: verdict-logic section (closest road, 30 ft rule, boundary edge)",
  src.includes('id="verdict"') && src.includes("Second road close") && src.includes("Urban boundary edge")
  && src.includes("closest road decides red vs blue")]);

let fail = 0;
for (const [label, ok] of checks) { console.log((ok ? "  ok   " : "  FAIL ") + label); if (!ok) fail++; }
if (errors.length) { console.log("page errors:"); errors.forEach(e => console.log("  " + e)); fail++; }
console.log(fail ? "VERIFY FAILED" : "VERIFY PASSED");
await browser.close();
process.exit(fail ? 1 : 0);
