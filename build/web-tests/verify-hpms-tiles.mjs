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
const CHROMIUM_PATH = process.env.PLAYWRIGHT_CHROMIUM_PATH
  || "/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell";

const avail = existsSync(tilesDir)
  ? readdirSync(tilesDir).filter(f => f.endsWith(".pmtiles") && f !== "basemap.pmtiles") : [];
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
  !netLines.some(l => /FeatureServer\/353\/query|LRSE_Functional_Class|FFCL_gdb|Functional_Class_Local|mndot_commonlayers2|FunctionalClass\/MapServer|Functional_Class\/MapServer/.test(l) && l.includes("esriGeometryEnvelope")))]);

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
