// Harvest the Region V ACUB polygons (2020 Adjusted Urban Areas) from the
// NTAD FeatureServer into NDJSON for tippecanoe — the cached-verdict path
// needs urban/rural offline, not just road classes.
//
//   node fetch-acub.mjs <out.ndjson>
//
// resultOffset pagination hits the same ~55 s server give-up the HPMS
// table does ("Unable to perform query", after 504s), so the harvest is
// two-phase: returnIdsOnly per state envelope (cheap, no geometry), then
// geometry in small objectIds batches with ~3 m server-side
// generalization (full precision also 504s; the verdict's boundary rule
// has a 76 m floor, so ≤3 m of boundary shift is noise).
const BASE = "https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/NTAD_Adjusted_Urban_Areas/FeatureServer/0";
const STATE_BOX = {
  26: [-90.5, 41.6, -82.0, 48.4], 55: [-93.0, 42.4, -86.1, 47.4],
  27: [-97.3, 43.4, -89.4, 49.5], 17: [-91.6, 36.9, -87.0, 42.6],
  18: [-88.2, 37.7, -84.7, 41.8], 39: [-84.9, 38.3, -80.4, 42.0],
};
import { createWriteStream } from "node:fs";

const outPath = process.argv[2];
if (!outPath) { console.error("usage: node fetch-acub.mjs <out.ndjson>"); process.exit(2); }

async function getJson(url, tries = 5) {
  for (let i = 0; ; i++) {
    try {
      const r = await fetch(url, { signal: AbortSignal.timeout(180000) });
      if (!r.ok) throw new Error("HTTP " + r.status);
      const j = await r.json();
      if (j.error) throw new Error("service error " + JSON.stringify(j.error).slice(0, 160));
      return j;
    } catch (e) {
      if (i >= tries - 1) throw e;
      console.log("  retry after:", e.message);
      await new Promise(res => setTimeout(res, 3000 * (i + 1)));
    }
  }
}

const ids = new Set();
for (const [fips, [x0, y0, x1, y1]] of Object.entries(STATE_BOX)) {
  const j = await getJson(`${BASE}/query?where=1%3D1&geometry=${x0},${y0},${x1},${y1}` +
    `&geometryType=esriGeometryEnvelope&inSR=4326&spatialRel=esriSpatialRelIntersects` +
    `&returnIdsOnly=true&f=json`);
  for (const id of j.objectIds || []) ids.add(id);
  console.log(`FIPS ${fips}: ${ (j.objectIds || []).length } ids (union ${ids.size})`);
}

const all = [...ids].sort((a, b) => a - b);
const out = createWriteStream(outPath);
let written = 0;
const BATCH = 25;
for (let i = 0; i < all.length; i += BATCH) {
  const j = await getJson(`${BASE}/query?objectIds=${all.slice(i, i + BATCH).join(",")}` +
    `&outFields=NAME,UACE,state_1&returnGeometry=true&outSR=4326&geometryPrecision=5` +
    `&maxAllowableOffset=0.00003&f=geojson`);
  for (const f of j.features || []) {
    if (!f.geometry) continue;
    out.write(JSON.stringify({ type: "Feature",
      properties: { NAME: f.properties.NAME, UACE: f.properties.UACE, state_1: f.properties.state_1 },
      geometry: f.geometry }) + "\n");
    written++;
  }
  console.log(`${Math.min(i + BATCH, all.length)}/${all.length} fetched (${written} written)`);
}
await new Promise(r => out.end(r));
console.log(`DONE: ${written} urban-area polygons -> ${outPath}`);
