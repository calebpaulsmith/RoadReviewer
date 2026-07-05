// Verifies the "Download PDF Report" feature (web/index.html's rr-report
// script block): classify a point, click the button, and check the
// downloaded PDF for a cover page, a per-site page with ONE combined
// page-filling figure (ACUB polygon + class polylines on one map), and the
// live source-map link annotations. Requires `npm install` in this
// directory first (playwright-core).
//
//   cd build/web-tests && npm install && node verify-pdf-report.mjs
//
// The query URLs and response shapes used by rr-report were independently
// confirmed live via curl against the real MDOT and NTAD ACUB services (see
// fixtures/*.json). mi-geom.json / acub-geom.json were captured live
// 2026-07-05 with rr-report's actual frame-ENVELOPE query shape (the 0.75 mi
// Kalamazoo frame: 88 road segments across 6 classes + the generalized
// Kalamazoo ACUB polygon), so the rendered figure in this test shows the
// realistic full street grid, not a stub. Chromium
// itself is stubbed with those real fixtures rather than hitting the network
// directly, since some sandboxes' outbound HTTPS proxy is only integrated
// with curl/Node's fetch, not with a standalone launched Chromium - if that's
// not true in your environment, feel free to delete the page.route() block
// below and let it hit the live services directly instead.
//
// Citations/legend/scale-bar are drawn ONTO each figure's canvas (so they
// travel with the raster image if it's later cropped out of the PDF), not
// emitted as separate PDF text objects - so this script checks PDF structure
// (page count, embedded image count) rather than grepping for citation text.
// Run with SAVE_SAMPLES=1 to also write sample-report.pdf and
// figure-sample.png next to this script for a visual spot-check.
import { chromium } from "playwright-core";
import { readFileSync, copyFileSync, writeFileSync } from "node:fs";
import { inflateSync } from "node:zlib";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const fixture = name => readFileSync(join(here, "fixtures", name), "utf8");
const miMeta = fixture("mi-meta.json"), miGeom = fixture("mi-geom.json");
const acubMeta = fixture("acub-meta.json"), acubGeom = fixture("acub-geom.json");

const PAGE = "file://" + join(here, "..", "..", "web", "index.html");
const CHROMIUM_PATH = process.env.PLAYWRIGHT_CHROMIUM_PATH
  || "/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell";

const browser = await chromium.launch({ executablePath: CHROMIUM_PATH });
const page = await browser.newPage({ viewport: { width: 1400, height: 950 } });

const errors = [];
page.on("pageerror", e => errors.push("pageerror: " + e.message));
page.on("console", m => {
  // The interactive Leaflet map's basemap tiles are deliberately route.abort()-ed
  // below (irrelevant to the report feature under test); ignore the resulting
  // "Failed to load resource" console noise they generate.
  if (m.type() === "error" && !m.text().includes("Failed to load resource")) errors.push("console: " + m.text());
});

await page.route("**/*", async route => {
  const url = route.request().url();
  if (url.startsWith("file://")) return route.continue();
  // World_Street_Map tiles: serve real captured tiles (fixtures/tiles/) with a
  // CORS header, exercising the taint-free basemap compositing path — a
  // tainted canvas would make toDataURL throw and fail the whole test. Tiles
  // outside the captured set (e.g. the interactive Leaflet map's low zooms)
  // and the imagery/labels layers abort as before.
  const tileM = url.match(/World_Street_Map\/MapServer\/tile\/(\d+)\/(\d+)\/(\d+)/);
  if (tileM) {
    try {
      const body = readFileSync(join(here, "fixtures", "tiles", `${tileM[1]}-${tileM[2]}-${tileM[3]}.jpg`));
      return route.fulfill({ contentType: "image/jpeg", headers: { "Access-Control-Allow-Origin": "*" }, body });
    } catch { return route.abort(); }
  }
  if (url.includes("arcgisonline.com")) return route.abort();   // imagery/labels layers: irrelevant to this test

  const json = { contentType: "application/json" };
  // --- classification pass (rr-core, returnGeometry=false) ---
  if (url.includes("FeatureServer/353/query") && !url.includes("returnGeometry=true"))
    return route.fulfill({ ...json, body: JSON.stringify({ features: [{ attributes: { FunctionalSystem: 6, PR: "0006904" } }] }) });
  if (url.includes("FeatureServer/543/query")) return route.fulfill({ ...json, body: JSON.stringify({ features: [] }) });
  if (url.includes("NTAD_Adjusted_Urban_Areas/FeatureServer/0/query") && !url.includes("returnGeometry=true"))
    return route.fulfill({ ...json, body: JSON.stringify({ features: [{ attributes: { NAME: "Kalamazoo, MI", UACE: "43723", state_1: "MI" } }] }) });
  if (url.includes("TIGERweb")) return route.fulfill({ ...json, body: JSON.stringify({ features: [{ attributes: { NAME: "S Pitcher St" } }] }) });

  // --- report pass (rr-report, returnGeometry=true + layer metadata) ---
  if (url.includes("FeatureServer/353?f=json")) return route.fulfill({ ...json, body: miMeta });
  if (url.includes("FeatureServer/353/query") && url.includes("returnGeometry=true")) return route.fulfill({ ...json, body: miGeom });
  if (url.includes("NTAD_Adjusted_Urban_Areas/FeatureServer/0?f=json")) return route.fulfill({ ...json, body: acubMeta });
  if (url.includes("NTAD_Adjusted_Urban_Areas/FeatureServer/0/query") && url.includes("returnGeometry=true")) return route.fulfill({ ...json, body: acubGeom });

  return route.abort();
});

await page.goto(PAGE, { waitUntil: "domcontentloaded" });
await page.fill("#coordsIn", "Kalamazoo culvert,42.28536,-85.57025");
await page.waitForFunction(() => (document.getElementById("statusCount").textContent || "").includes("1 point(s) classified"), { timeout: 15000 });
const verdictRowOk = (await page.locator("#resultsBody .row.v-fed").first().textContent())?.includes("Kalamazoo");

const [download] = await Promise.all([
  page.waitForEvent("download", { timeout: 30000 }),
  page.click("#pdfBtn"),
]);
const dlPath = await download.path();
const bytes = readFileSync(dlPath);
// The report is generated with compress:true, so page-content streams are
// Flate-encoded — inflate every stream and append the decoded text so the
// string checks below can still grep for it. Link annotations and object
// dictionaries stay uncompressed and are covered by the raw latin1 text.
let text = bytes.toString("latin1");
{
  const raw = bytes.toString("latin1");
  const re = /stream\r?\n/g;
  let m;
  while ((m = re.exec(raw)) !== null) {
    const start = m.index + m[0].length;
    const end = raw.indexOf("endstream", start);
    if (end < 0) continue;
    try { text += "\n" + inflateSync(bytes.subarray(start, end)).toString("latin1"); } catch { /* image data etc. */ }
  }
}

// Direct probe of the basemap path: the default Kalamazoo frame at z16 needs
// exactly the 12 captured tiles, all of which must load via crossOrigin
// without tainting.
const basemapTileCount = await page.evaluate(async () => {
  const f = reportFrame(42.28536, -85.57025, 600);
  const bm = await fetchBasemapTiles(f, 748);
  return bm ? bm.tiles.length : 0;
});
const netLogText = await page.evaluate(() => document.getElementById("netLog").textContent);

const checks = [
  ["classification produced the expected federal-aid row", verdictRowOk],
  ["street basemap tiles load for the frame (12 at z16)", basemapTileCount === 12],
  ["basemap tile fetch disclosed in the network log", netLogText.includes("basemap tiles")],
  ["PDF magic bytes", bytes.slice(0, 5).toString("ascii") === "%PDF-"],
  ["non-trivial size (>50KB, images embedded)", bytes.length > 50_000],
  ["cover page title", text.includes("Federal-Aid Classification Report")],
  ["site name in cover table", text.includes("Kalamazoo culvert")],
  ["verdict in cover table", text.includes("Federal aid - Urban Minor Collector")],
  ["2 pages (cover + 1 site)", (text.match(/\/Type\s*\/Page[^s]/g) || []).length === 2],
  ["2 image XObjects embedded (1 combined figure x RGB+alpha)", (text.match(/\/Subtype\s*\/Image/g) || []).length === 2],
  // MI carries BOTH the official-app primary reference (canonical root URL,
  // no fragile hash) AND the restored first-tier pinned FEMA-viewer webmap.
  ["MI official-app reference present (canonical root, no hash)", text.includes("experience.arcgis.com/experience/7edd160c205d46b481fcd605bb4c58ce") && !text.includes("widget_167")],
  ["MI first-tier pinned FEMA webmap link present", text.includes("webmap=6a1702b9147243d1a5ee62cd614bc681")],
  ["Google Maps link annotation present", text.includes("google.com/maps?q=42.28536")],
  ["zoom select defaults to Standard (600 m half-width)", (await page.locator("#pdfZoom").inputValue()) === "600"],
  ["disclaimer present", text.includes("classifies the road, not the project")],
  ["button label restored after run", (await page.locator("#pdfBtn").textContent()) === "Download PDF Report"],
];

let fail = 0;
for (const [label, ok] of checks) { console.log((ok ? "  ok   " : "  FAIL ") + label); if (!ok) fail++; }
if (errors.length) { console.log("page errors:"); errors.forEach(e => console.log("  " + e)); fail++; }
console.log(fail ? "VERIFY FAILED" : "VERIFY PASSED", "| pdf bytes:", bytes.length);

if (process.env.SAVE_SAMPLES) {
  copyFileSync(dlPath, join(here, "sample-report.pdf"));
  const figurePng = await page.evaluate(async ({ miMetaStr, miGeomStr, acubMetaStr, acubGeomStr }) => {
    const miM = JSON.parse(miMetaStr), miG = JSON.parse(miGeomStr);
    const acM = JSON.parse(acubMetaStr), acG = JSON.parse(acubGeomStr);
    const point = { lat: 42.28536, lon: -85.57025 };
    const frame = reportFrame(point.lat, point.lon, 600);
    const basemap = await fetchBasemapTiles(frame, 748);
    return renderCombinedCanvas({
      title: "MI road functional class + 2020 Adjusted Urban Boundary",
      layers: [
        { geometryType: "polygon", drawingInfo: acM.drawingInfo, features: acG.features,
          legendHeader: "Urban boundary (USDOT NTAD 2020)" },
        { geometryType: "polyline", rendererField: "FunctionalSystem", drawingInfo: miM.drawingInfo,
          features: miG.features, legendHeader: "Road functional class (MDOT)" },
      ],
      point, verdictColor: "#d73027", frame, basemap,
      citationLines: ["Road class: Functional System — https://mdotgis.state.mi.us/.../FeatureServer/353",
        "Urban boundary: USDOT NTAD 2020 Adjusted Urban Area Boundaries — https://services.arcgis.com/...",
        "Basemap: © Esri World Street Map tiles, fetched for this frame at report time",
        "Retrieved (test) · frame ≈ 0.75 mi wide · classification buffer 200 ft"],
      notes: [],
    }).toDataURL("image/png");
  }, { miMetaStr: miMeta, miGeomStr: miGeom, acubMetaStr: acubMeta, acubGeomStr: acubGeom });
  writeFileSync(join(here, "figure-sample.png"), Buffer.from(figurePng.split(",")[1], "base64"));
  console.log("Saved sample-report.pdf and figure-sample.png next to this script.");
}

await browser.close();
process.exit(fail ? 1 : 0);
