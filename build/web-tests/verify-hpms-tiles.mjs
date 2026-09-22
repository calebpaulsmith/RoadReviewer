// Verifies the baked-HPMS-tiles class display in web/index.html against a
// REAL built tileset (web/tiles/<st>.pmtiles must exist — build it with
// build/tiles/build-state-tiles.sh first; the script skips with a clear
// message if none is present).
//
// PMTiles is read via HTTP range requests, so the page is served by a
// tiny range-capable static server here (python http.server can't do
// ranges). State/ACUB/classification services are stubbed with the same
// fixtures as verify-review-ui; the TILE fetches are real.
//
// Asserts:
//   - the legend switches to the "FHWA HPMS … tiles" section for the state
//   - protomaps-leaflet drew into the browse pane with real colored pixels
//   - NO live road-class envelope query fired for the tile-served state
//
//   cd build/web-tests && node verify-hpms-tiles.mjs
import { chromium } from "playwright-core";
import { createServer } from "node:http";
import { createReadStream, readFileSync, statSync, existsSync, readdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join, normalize } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const webRoot = join(here, "..", "..", "web");
const tilesDir = join(webRoot, "tiles");
// /opt/pw-browsers/chromium is a stable symlink to the installed build; a
// version-pinned path (or playwright's own bundled default) rots whenever the
// sandbox image bumps, and fails with a misleading "npx playwright install".
const CHROMIUM_PATH = process.env.PLAYWRIGHT_CHROMIUM_PATH || "/opt/pw-browsers/chromium";

const avail = existsSync(tilesDir)
  ? readdirSync(tilesDir).filter(f => f.endsWith(".pmtiles") && f !== "basemap.pmtiles" && f !== "acub.pmtiles") : [];
if (!avail.length) {
  console.log("SKIP: no web/tiles/*.pmtiles present — build one with build/tiles/build-state-tiles.sh");
  process.exit(0);
}
// Test with whichever state tileset exists; view centers come from here.
const CENTER = { mi: [42.29, -85.58], wi: [43.04, -87.91], in: [39.77, -86.16],
                 mn: [44.95, -93.17], il: [41.90, -87.69], oh: [40.01, -82.99] };
const st = avail[0].replace(".pmtiles", "");
const [lat, lon] = CENTER[st] || [42.29, -85.58];
console.log(`testing with web/tiles/${st}.pmtiles`);

// ---- range-capable static server over web/ ----
const server = createServer((req, res) => {
  const path = normalize(join(webRoot, decodeURIComponent(req.url.split("?")[0])));
  if (!path.startsWith(webRoot)) { res.writeHead(403); return res.end(); }
  let st_;
  try { st_ = statSync(path); } catch { res.writeHead(404); return res.end(); }
  const type = path.endsWith(".html") ? "text/html" : path.endsWith(".js") ? "text/javascript"
    : path.endsWith(".css") ? "text/css" : "application/octet-stream";
  const range = req.headers.range && /^bytes=(\d+)-(\d+)?$/.exec(req.headers.range);
  if (req.method === "HEAD") { res.writeHead(200, { "Content-Type": type, "Content-Length": st_.size, "Accept-Ranges": "bytes" }); return res.end(); }
  if (range) {
    const start = parseInt(range[1], 10);
    const end = range[2] ? Math.min(parseInt(range[2], 10), st_.size - 1) : st_.size - 1;
    res.writeHead(206, { "Content-Type": type, "Accept-Ranges": "bytes",
      "Content-Range": `bytes ${start}-${end}/${st_.size}`, "Content-Length": end - start + 1 });
    return createReadStream(path, { start, end }).pipe(res);
  }
  res.writeHead(200, { "Content-Type": type, "Content-Length": st_.size, "Accept-Ranges": "bytes" });
  createReadStream(path, { start: 0 }).pipe(res);
});
await new Promise(res => server.listen(0, "127.0.0.1", res));
const port = server.address().port;

const fixture = name => readFileSync(join(here, "fixtures", name), "utf8");
const acubMeta = fixture("acub-meta.json"), acubGeom = fixture("acub-geom.json");

const browser = await chromium.launch({ executablePath: CHROMIUM_PATH });
const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
const errors = [], reqUrls = [];
page.on("pageerror", e => errors.push("pageerror: " + e.message));
page.on("request", r => reqUrls.push(r.url()));

await page.route("**/*", async route => {
  const url = route.request().url();
  if (url.includes(`127.0.0.1:${port}`)) return route.continue();   // page + REAL tiles
  const json = { contentType: "application/json" };
  if (url.includes("arcgisonline.com")) return route.abort();
  if (url.includes("NTAD_Adjusted_Urban_Areas/FeatureServer/0?f=json")) return route.fulfill({ ...json, body: acubMeta });
  if (url.includes("NTAD_Adjusted_Urban_Areas/FeatureServer/0/query")) return route.fulfill({ ...json, body: acubGeom });
  return route.abort();   // everything else (incl. state road layers) blocked — tiles must carry the display
});

await page.goto(`http://127.0.0.1:${port}/index.html`, { waitUntil: "domcontentloaded" });
await page.evaluate(([la, lo]) => map.setView([la, lo], 13), [lat, lon]);

const checks = [];
await page.waitForFunction(() => {
  const el = document.getElementById("liveLegend");
  return el && el.textContent.includes("FHWA HPMS") && el.textContent.includes("tiles");
}, { timeout: 20000 });
checks.push(["legend switches to the baked HPMS tiles section", true]);
checks.push(["legend lists all 7 FHWA classes", await page.evaluate(() => {
  const t = document.getElementById("liveLegend").textContent;
  return t.includes("Interstate") && t.includes("Local") && t.includes("verdicts") || t.includes("Verdicts");
})]);
// Source provenance (2026-09-17): every class row in the HPMS legend section
// is a link to FHWA's public HPMS page with a hover title naming the source.
checks.push(["HPMS legend rows link to FHWA's public HPMS page with a source tooltip", await page.evaluate(() => {
  const rows = [...document.querySelectorAll("#liveLegend a.lrow")].filter(a => /fhwa\.dot\.gov/.test(a.href));
  return rows.length >= 7 && rows.every(a => /HPMS/.test(a.title) && a.target === "_blank" && !/\/rest\//.test(a.href));
})]);

// tiles actually painted: wait for a browse-pane canvas with non-blank pixels
await page.waitForFunction(() => {
  for (const c of document.querySelectorAll(".leaflet-browse-pane canvas")) {
    try {
      const ctx = c.getContext("2d");
      const d = ctx.getImageData(0, 0, c.width, c.height).data;
      let n = 0;
      for (let i = 3; i < d.length; i += 40) if (d[i] > 0) n++;
      if (n > 50) return true;
    } catch { /* keep looking */ }
  }
  return false;
}, { timeout: 30000 });
checks.push(["protomaps-leaflet painted real tile pixels in the browse pane", true]);

checks.push(["no live road-class query fired for the tile-served state", await page.evaluate(() =>
  !netLines.some(l => /MapServer\/353\/query|LRSE_Functional_Class|FFCL_gdb|Functional_Class_Local|mndot_commonlayers2|FunctionalClass\/MapServer|Functional_Class\/MapServer/.test(l) && l.includes("esriGeometryEnvelope")))]);

// --- class lines OVERZOOM past the tileset's z13 (maxDataZoom regression) ---
await page.evaluate(([la, lo]) => {
  for (const c of document.querySelectorAll(".leaflet-browse-pane canvas"))
    c.getContext("2d").clearRect(0, 0, c.width, c.height);   // don't let stale z13 canvases mask a blank z15
  map.setView([la, lo], 15);
}, [lat, lon]);
await page.waitForFunction(() => {
  for (const c of document.querySelectorAll(".leaflet-browse-pane canvas")) {
    try {
      const d = c.getContext("2d").getImageData(0, 0, c.width, c.height).data;
      let n = 0;
      for (let i = 3; i < d.length; i += 40) if (d[i] > 0) n++;
      if (n > 50) return true;
    } catch { /* keep looking */ }
  }
  return false;
}, { timeout: 30000 });
checks.push(["class tiles overzoom past z13 (painted at z15)", true]);
await page.evaluate(([la, lo]) => map.setView([la, lo], 13), [lat, lon]);

// --- cached-tile CLASSIFICATION: verdict from hosted tiles, no live class/ACUB query ---
if (existsSync(join(tilesDir, "acub.pmtiles"))) {
  // The state's live-verified "Federal aid" test point (§4.2/§4.2a-e #1)
  const TESTPT = { mi: [42.28536, -85.57025], in: [39.7684, -86.1581], wi: [43.0389, -87.9065],
                   mn: [44.9531, -93.1668], il: [41.9020, -87.6870], oh: [40.0150, -82.9990] };
  const [tLat, tLon] = TESTPT[st] || TESTPT.mi;
  const before = reqUrls.length;
  await page.fill("#coordsIn", `Cached check,${tLat},${tLon}`);
  await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("1 point(s) classified"), { timeout: 30000 });
  const rowText = await page.evaluate(() => document.querySelector("#resultsBody .row").textContent);
  const cachedOk = /federal aid/i.test(rowText) && rowText.includes("Source: FHWA HPMS")
    && rowText.includes("State route") && !rowText.includes("Failed");   // linkage chip from the R/B/E attrs
  if (!cachedOk) console.log("  row text was:", rowText.slice(0, 300));
  checks.push(["cached classification returns the known Federal-aid verdict from the tiles", cachedOk]);
  checks.push(["row Source chip links to FHWA's public HPMS page (not a REST URL)", await page.evaluate(() => {
    const a = document.querySelector("#resultsBody .row a.srcchip");
    return !!a && /fhwa\.dot\.gov/.test(a.href) && !/\/rest\//.test(a.href) && /Data source/.test(a.title);
  })]);
  checks.push(["exports carry a Data Source column naming the HPMS tiles", await page.evaluate(() =>
    EXPORT_HEADERS.includes("Data Source") && exportRows().every(r => /FHWA HPMS/.test(r[EX.source])))]);

  checks.push(["no live class/ACUB point query fired for the cached verdict",
    !reqUrls.slice(before).some(u =>
      /MapServer\/353\/query|LRSE_Functional_Class|FFCL_gdb|Functional_Class_Local|mndot_commonlayers2|FunctionalClass\/MapServer|Functional_Class\/MapServer|NTAD_Adjusted_Urban_Areas/.test(u)
      && u.includes("esriGeometryPoint"))]);
  // --- source outage -> HPMS fallback (2026-09-17). Live verdicts ON, the
  // state's class server answering an HTML 500: the verdict still comes
  // (from the tiles), the row + export say which source was down, the
  // strip reports it, the down host is never queried again, and a JSON
  // probe brings it back with a re-run offer. ---
  const ST = st.toUpperCase();
  const stHost = await page.evaluate(S => new URL(svc(STATE_SVC_KEY[S])).host, ST);
  const dot = await page.evaluate(S => DOT_NAME[S], ST);
  // health is tracked per SERVICE (host + ArcGIS service path), not per host
  const stKey = await page.evaluate(S => hostOf(svc(STATE_SVC_KEY[S])), ST);
  const stUrl = await page.evaluate(S => svc(STATE_SVC_KEY[S]), ST);
  await page.route(`**/${stHost}/**`, route => route.fulfill({ status: 500, contentType: "text/html", body: "<html>500</html>" }));
  await page.evaluate(() => { document.getElementById("liveVerdicts").checked = true; });
  await page.fill("#coordsIn", `Outage check,${tLat + 0.0002},${tLon}`);
  await page.waitForFunction(() => { const a = document.querySelector("#resultsBody .row a.srcchip"); return !!a && /down/.test(a.textContent); }, { timeout: 30000 });
  checks.push([`outage: live verdict falls back to the HPMS tiles and names ${dot} as down on the row`, await page.evaluate(d => {
    const a = document.querySelector("#resultsBody .row a.srcchip"), r = document.querySelector("#resultsBody .row");
    return !!a && /FHWA HPMS/.test(a.textContent) && a.textContent.includes(`${d} live layer down`) && /FALLBACK/.test(a.title)
      && !r.className.includes("v-failed") && /federal aid/i.test(r.textContent);
  }, dot)]);
  checks.push(["outage: status strip reports the source down (HTTP 500) with a recheck", await page.evaluate(d => {
    const s = document.getElementById("sourceStatus");
    return !s.hidden && s.textContent.includes(`${d} live layer`) && /HTTP 500/.test(s.textContent) && /rechecked/.test(s.textContent);
  }, dot)]);
  checks.push(["outage: exports' Data Source column carries the fallback note", await page.evaluate(d =>
    exportRows().every(r => /FHWA HPMS/.test(r[EX.source]) && r[EX.source].includes(`${d} live layer down`)), dot)]);
  checks.push(["outage: the down source is not queried again (fast-fail, no request)", await page.evaluate(async ([h, u]) => {
    const before = netLines.filter(l => l.includes("GET ") && l.includes(h)).length;
    try { await httpGetJson(u + "/query?f=json"); return false; } catch (e) { if (!e.sourceDown) return false; }
    return netLines.filter(l => l.includes("GET ") && l.includes(h)).length === before;
  }, [stHost, stUrl])]);
  await page.route(`**/${stHost}/**`, route => route.fulfill({ status: 200, contentType: "application/json", body: JSON.stringify({ name: "Functional Class" }) }));
  await page.evaluate(h => probeSource(h), stKey);
  await page.waitForFunction(() => /back up/.test(document.getElementById("sourceStatus").textContent), { timeout: 10000 });
  checks.push(["outage: recovered source reported back up with a re-run offer for the fallback row", await page.evaluate(() =>
    /re-run 1 row/.test(document.getElementById("sourceStatus").textContent))]);
  await page.evaluate(() => { document.getElementById("liveVerdicts").checked = false; });
} else {
  console.log("  (skip) web/tiles/acub.pmtiles not present — cached-classification check skipped");
}

// --- offline road basemap (web/tiles/basemap.pmtiles, Protomaps extract) ---
if (existsSync(join(tilesDir, "basemap.pmtiles"))) {
  await page.waitForFunction(() => {
    for (const c of document.querySelectorAll(".leaflet-tile-pane canvas")) {
      try {
        const d = c.getContext("2d").getImageData(0, 0, c.width, c.height).data;
        let n = 0;
        for (let i = 3; i < d.length; i += 40) if (d[i] > 0) n++;
        if (n > 50) return true;
      } catch { /* keep looking */ }
    }
    return false;
  }, { timeout: 30000 });
  checks.push(["offline Protomaps road basemap painted (no Esri street-tile requests)",
    !reqUrls.some(u => u.includes("World_Street_Map"))]);
} else {
  console.log("  (skip) web/tiles/basemap.pmtiles not present — offline basemap check skipped");
}

let failed = 0;
for (const [name, ok] of checks) { console.log((ok ? "  ok   " : "  FAIL ") + name); if (!ok) failed++; }
const realErrors = errors.filter(e => !e.includes("Failed to load resource"));
if (realErrors.length) { console.log("page errors:", realErrors.join(" | ")); failed++; }
await browser.close();
server.close();
console.log(failed ? "VERIFY FAILED" : "VERIFY PASSED");
process.exit(failed ? 1 : 0);
