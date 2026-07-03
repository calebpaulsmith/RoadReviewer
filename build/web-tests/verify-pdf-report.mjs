// Verifies the "Download PDF Report" feature (web/index.html's rr-report
// script block): classify a point, click the button, and check the
// downloaded PDF for a cover page, per-site pages, and embedded figure
// images. Requires `npm install` in this directory first (playwright-core).
//
//   cd build/web-tests && npm install && node verify-pdf-report.mjs
//
// The query URLs and response shapes used by rr-report were independently
// confirmed live via curl against the real MDOT and NTAD ACUB services (see
// fixtures/*.json, captured from those live responses - MI query used
// FunctionalSystem/PR outFields with a 200ft fallback buffer; ACUB used
// NAME/UACE/state_1 with maxAllowableOffset=0.0001 generalization). Chromium
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
  if (url.includes("arcgisonline.com")) return route.abort();   // basemap tiles: irrelevant to this test

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
const verdictRowOk = (await page.locator("tr.v-fed td").first().textContent())?.includes("Kalamazoo");

const [download] = await Promise.all([
  page.waitForEvent("download", { timeout: 30000 }),
  page.click("#pdfBtn"),
]);
const dlPath = await download.path();
const bytes = readFileSync(dlPath);
const text = bytes.toString("latin1");

const checks = [
  ["classification produced the expected federal-aid row", verdictRowOk],
  ["PDF magic bytes", bytes.slice(0, 5).toString("ascii") === "%PDF-"],
  ["non-trivial size (>50KB, images embedded)", bytes.length > 50_000],
  ["cover page title", text.includes("Federal-Aid Classification Report")],
  ["site name in cover table", text.includes("Kalamazoo culvert")],
  ["verdict in cover table", text.includes("Federal aid - Urban Minor Collector")],
  ["2 pages (cover + 1 site)", (text.match(/\/Type\s*\/Page[^s]/g) || []).length === 2],
  ["4 image XObjects embedded (2 figures x RGB+alpha)", (text.match(/\/Subtype\s*\/Image/g) || []).length === 4],
  ["disclaimer present", text.includes("classifies the road, not the project")],
  ["button label restored after run", (await page.locator("#pdfBtn").textContent()) === "Download PDF Report"],
];

let fail = 0;
for (const [label, ok] of checks) { console.log((ok ? "  ok   " : "  FAIL ") + label); if (!ok) fail++; }
if (errors.length) { console.log("page errors:"); errors.forEach(e => console.log("  " + e)); fail++; }
console.log(fail ? "VERIFY FAILED" : "VERIFY PASSED", "| pdf bytes:", bytes.length);

if (process.env.SAVE_SAMPLES) {
  copyFileSync(dlPath, join(here, "sample-report.pdf"));
  const figurePng = await page.evaluate(({ miMetaStr, miGeomStr }) => {
    const meta = JSON.parse(miMetaStr), geom = JSON.parse(miGeomStr);
    return renderFigureCanvas({
      title: "MI road functional class — Functional System",
      features: geom.features, geometryType: "polyline", rendererField: "FunctionalSystem",
      drawingInfo: meta.drawingInfo, point: { lat: 42.28536, lon: -85.57025 },
      verdictColor: "#d73027", radiusMeters: 250,
      citationLines: ["Source: Functional System (Esri REST feature service)", "https://mdotgis.state.mi.us/.../FeatureServer/353", "Retrieved (test) · query buffer 200 ft"],
      emptyNote: "No road segment returned within 200 ft",
    }).toDataURL("image/png");
  }, { miMetaStr: miMeta, miGeomStr: miGeom });
  writeFileSync(join(here, "figure-sample.png"), Buffer.from(figurePng.split(",")[1], "base64"));
  console.log("Saved sample-report.pdf and figure-sample.png next to this script.");
}

await browser.close();
process.exit(fail ? 1 : 0);
