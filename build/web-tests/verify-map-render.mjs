// Verifies the interactive map's RENDERING against the real tilesets —
// the four bugs reported 2026-09-14 (blank grey land at the region view,
// river "slivers" from a v3 renderer on v4 tiles, the region cut to a
// rectangle instead of the state shapes, faint/invisible local class
// lines at street zoom) plus a main-thread budget while zooming.
//
// Needs web/tiles/basemap.pmtiles, web/tiles/acub.pmtiles and
// web/tiles/mi.pmtiles (the checks are pixel reads against the real
// files, served by the same range-capable static server as
// verify-hpms-tiles). OpenFreeMap (street-level detail) is fetched live
// and is best-effort here: its absence does not fail the run.
//
// Asserts:
//   - region view (map minZoom): the basemap paints roads/labels/borders
//     (many distinct colours, not just land + water), NO class lines yet
//     (the opening map is a clean atlas view; classes start at z9), and a
//     point outside the six states shows the mask ground colour
//   - Minneapolis z10: the ACUB tiles paint pink, and no cyan "sliver" pixel
//     sits on a land point that the old renderer filled
//   - Long Lake, MI z17: the class-7 subdivision street (Totem Trl) has
//     opaque dark pixels in the MI class layer's tile (width scaling)
//   - a 6->18 zoom sequence over Milwaukee produces no single long task
//     over 1.5 s and no page errors
//
//   cd build/web-tests && node verify-map-render.mjs
import { chromium } from "playwright-core";
import { createServer } from "node:http";
import { createReadStream, statSync, existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join, normalize } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const webRoot = join(here, "..", "..", "web");
// /opt/pw-browsers/chromium is a stable symlink to the installed build; a
// version-pinned path (or playwright's own bundled default) rots whenever the
// sandbox image bumps, and fails with a misleading "npx playwright install".
const CHROMIUM_PATH = process.env.PLAYWRIGHT_CHROMIUM_PATH || "/opt/pw-browsers/chromium";
for (const f of ["tiles/basemap.pmtiles", "tiles/acub.pmtiles", "tiles/mi.pmtiles", "data/r5-states.geojson"])
  if (!existsSync(join(webRoot, f))) { console.log(`SKIP: web/${f} missing`); process.exit(0); }

const TYPES = { ".html": "text/html", ".js": "application/javascript", ".css": "text/css", ".pmtiles": "application/octet-stream",
                ".png": "image/png", ".json": "application/json", ".geojson": "application/geo+json" };
const server = createServer((req, res) => {
  let p = decodeURIComponent(req.url.split("?")[0]); if (p === "/") p = "/index.html";
  const file = normalize(join(webRoot, p));
  if (!file.startsWith(normalize(webRoot)) || !existsSync(file) || statSync(file).isDirectory()) { res.writeHead(404); return res.end(); }
  const st = statSync(file), type = TYPES[file.slice(file.lastIndexOf("."))] || "application/octet-stream";
  const range = req.headers.range && /^bytes=(\d+)-(\d+)?$/.exec(req.headers.range);
  if (req.method === "HEAD") { res.writeHead(200, { "Content-Type": type, "Content-Length": st.size, "Accept-Ranges": "bytes" }); return res.end(); }
  if (range) {
    const s = +range[1], e = range[2] ? Math.min(+range[2], st.size - 1) : st.size - 1;
    res.writeHead(206, { "Content-Type": type, "Accept-Ranges": "bytes", "Content-Range": `bytes ${s}-${e}/${st.size}`, "Content-Length": e - s + 1 });
    return createReadStream(file, { start: s, end: e }).pipe(res);
  }
  res.writeHead(200, { "Content-Type": type, "Content-Length": st.size, "Accept-Ranges": "bytes" });
  createReadStream(file).pipe(res);
});
await new Promise(r => server.listen(0, "127.0.0.1", r));
const port = server.address().port;

let failures = 0;
const ok = (cond, label, detail = "") => {
  console.log(`  ${cond ? "ok  " : "FAIL"} ${label}${detail ? " - " + detail : ""}`);
  if (!cond) failures++;
};
const browser = await chromium.launch({ executablePath: CHROMIUM_PATH });
const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
const pageErrors = [];
page.on("pageerror", e => pageErrors.push(e.message));
await page.goto(`http://127.0.0.1:${port}/index.html`, { waitUntil: "load" });
await page.evaluate(() => {
  window.__lt = [];
  new PerformanceObserver(l => { for (const e of l.getEntries()) window.__lt.push(Math.round(e.duration)); }).observe({ type: "longtask", buffered: true });
});
const sleep = ms => page.waitForTimeout(ms);

// Pixel helpers run in-page: read a pane's canvases at a container point,
// or sample colours across a whole pane.
const paneAt = (pane, lat, lon) => page.evaluate(([pane, lat, lon]) => {
  const pt = map.latLngToContainerPoint([lat, lon]);
  const mr = document.getElementById("map").getBoundingClientRect();
  const x = mr.left + pt.x, y = mr.top + pt.y; const out = [];
  for (const c of document.querySelector("." + pane).querySelectorAll("canvas")) {
    const r = c.getBoundingClientRect();
    if (x >= r.left && x < r.right && y >= r.top && y < r.bottom)
      out.push([...c.getContext("2d").getImageData(Math.round((x - r.left) * c.width / r.width), Math.round((y - r.top) * c.height / r.height), 1, 1).data]);
  }
  return out;
}, [pane, lat, lon]);
const paneColours = (pane, minAlpha = 200) => page.evaluate(([pane, minAlpha]) => {
  const set = new Set(); let blue = 0;
  for (const c of document.querySelector("." + pane).querySelectorAll("canvas")) {
    const d = c.getContext("2d").getImageData(0, 0, c.width, c.height).data;
    for (let i = 0; i < d.length; i += 16) {
      if (d[i + 3] < minAlpha) continue;
      set.add(((d[i] >> 4) << 8) | ((d[i + 1] >> 4) << 4) | (d[i + 2] >> 4));
      if (d[i] < 40 && d[i + 1] > 90 && d[i + 1] < 140 && d[i + 2] > 220) blue++;
    }
  }
  return { distinct: set.size, blue };
}, [pane, minAlpha]);

// ---- 1. region view ----
await sleep(9000);
const z0 = await page.evaluate(() => map.getZoom());
const base = await paneColours("leaflet-tile-pane");
ok(base.distinct > 40, "region view: basemap paints more than land + water", `${base.distinct} distinct colours at z${z0}`);
const browse = await paneColours("leaflet-browse-pane");
ok(browse.blue === 0, "region view: NO class lines at the region zoom (clean opening map)", `${browse.blue} sampled blue px`);
const outside = await paneAt("leaflet-mask-pane", 38.5, -95.5);        // Kansas: outside the six states
ok(outside.some(p => p[3] > 250), "region view: mask covers points outside the six states", JSON.stringify(outside[0]));
const inside = await paneAt("leaflet-mask-pane", 44.9, -93.2);         // Minneapolis: inside a hole
ok(inside.every(p => p[3] === 0), "region view: mask leaves the states open", JSON.stringify(inside[0]));

// ---- 2. Minneapolis z10: ACUB from tiles, no river slivers ----
await page.evaluate(() => map.setView([44.95, -93.5], 10)); await sleep(8000);
const browse10 = await paneColours("leaflet-browse-pane");
ok(browse10.blue > 200, "z10: class tiles paint interstate blue", `${browse10.blue} sampled blue px`);
const pink = await paneAt("leaflet-acub-pane", 44.98, -93.27);         // inside the Minneapolis urban area
ok(pink.some(p => p[3] > 30 && p[0] > 200 && p[2] > 150), "z10: ACUB tiles paint the urban area", JSON.stringify(pink[0]));
const land = await paneAt("leaflet-tile-pane", 44.760, -93.610);        // farmland SW of the metro; a sliver crossed here before
ok(land.every(p => !(p[0] < 160 && p[1] > 200 && p[2] > 200)), "z10: no cyan sliver on a land point", JSON.stringify(land[0]));

// ---- 3. Long Lake, MI z17: local class line visible ----
await page.evaluate(() => map.setView([41.8806, -84.7905], 17)); await sleep(9000);
const totem = await page.evaluate(() => {
  const e = hpmsTile.get("MI"); if (!e || !e.layer) return null;
  const z = map.getZoom(), p = map.project([41.8806, -84.79055], z);
  const t = e.layer._tiles[Math.floor(p.x / 256) + ":" + Math.floor(p.y / 256) + ":" + z]; if (!t) return null;
  const d = t.el.getContext("2d").getImageData(0, 0, 256, 256).data; let dark = 0;
  for (let i = 0; i < d.length; i += 4) if (d[i + 3] > 200 && d[i] < 90 && d[i + 1] < 90 && d[i + 2] < 90) dark++;
  return dark;
});
ok(totem !== null && totem > 150, "z17: class-7 Totem Trl paints opaque dark pixels in the MI class tile", `${totem} px`);
const attached = await page.evaluate(() => [...hpmsTile].filter(([s, e]) => e.layer && map.hasLayer(e.layer)).map(([s]) => s));
ok(attached.length <= 3 && attached.includes("MI"), "z17 in Michigan: only in-view state layers stay attached", attached.join(","));

// ---- 4. zoom sequence budget ----
await page.evaluate(() => { window.__lt = []; map.setView([43.0389, -87.9065], 6); });
for (let z = 6; z <= 18; z++) { await page.evaluate(z => map.setZoom(z), z); await sleep(150); }
await sleep(8000);
const lt = await page.evaluate(() => window.__lt);
ok(Math.max(0, ...lt) < 1500, "zoom 6->18 over Milwaukee: no long task over 1.5 s", `max ${Math.max(0, ...lt)} ms, total ${lt.reduce((a, b) => a + b, 0)} ms`);
ok(pageErrors.length === 0, "no page errors", pageErrors.join(" | "));

await browser.close(); server.close();
console.log(failures ? `VERIFY FAILED (${failures})` : "VERIFY PASSED");
process.exit(failures ? 1 : 0);
