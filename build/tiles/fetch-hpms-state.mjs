// Extract one state's full-extent road network (every public road, FHWA
// class 1-7) from the BTS-hosted nationwide HPMS layer into newline-
// delimited GeoJSON ready for tippecanoe.
//
//   node fetch-hpms-state.mjs <STATE_FIPS> <out.ndjson>
//
// Source: HPMS_National_Current (services.arcgis.com/xOi1kZaI0eWDREZv —
// the same USDOT/BTS AGOL org that hosts the NTAD ACUB layer). Confirmed
// live 2026-09-14: full extent incl. locals; STATE_ID = state FIPS;
// F_SYSTEM = FHWA 1-7; per-state counts WI 902,611 / IN 530,981 /
// MI 354,291 / MN 509,570 / IL 445,757 / OH 487,411.
//
// HARVEST STRATEGY — quadtree envelopes, not offset pagination. The
// national table is ~30M rows and any attribute-filtered offset page
// times out server-side (~55 s -> HTTP 400, measured live), while an
// envelope query answers in seconds because it rides the spatial index.
// So: start from the state's bounding box; any cell that returns the
// 2,000-feature cap with exceededTransferLimit splits into 4; leaf cells
// stream to the output. Segments crossing cell borders arrive twice and
// are deduped by OBJECTID.
//
// Each output feature carries:
//   properties.F  — FHWA functional class 1-7 (0/null dropped)
//   tippecanoe.minzoom — display band by class, so the finished tileset
//     itself implements progressive display: interstates z6, other
//     principal arterials z8, minor arterials z9, major collectors z10,
//     minor collectors z11, locals z12. tippecanoe honors the per-feature
//     "tippecanoe" key and drops it from the output tiles.

import { createWriteStream } from "node:fs";
import { mkdir } from "node:fs/promises";
import { dirname } from "node:path";

const BASE = "https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/HPMS_National_Current/FeatureServer/0";
const PAGE_CAP = 2000;        // the layer's maxRecordCount
const CONCURRENCY = 4;
const MIN_CELL_DEG = 0.005;   // never split below ~500 m — data-error guard
const MINZOOM_BY_CLASS = { 1: 6, 2: 6, 3: 8, 4: 9, 5: 10, 6: 11, 7: 12 };

// Generous per-state boxes (same rough bounds the web tool's detectState
// uses, padded) — the STATE_ID where-clause clips exactly.
const STATE_BOX = {
  26: [-90.5, 41.6, -82.0, 48.4],   // MI
  55: [-93.0, 42.4, -86.1, 47.4],   // WI
  27: [-97.3, 43.4, -89.4, 49.5],   // MN
  17: [-91.6, 36.9, -87.0, 42.6],   // IL
  18: [-88.2, 37.7, -84.7, 41.8],   // IN
  39: [-84.9, 38.3, -80.4, 42.0],   // OH
};

const [, , fipsArg, outPath] = process.argv;
const fips = parseInt(fipsArg, 10);
const box = STATE_BOX[fips];
if (!box || !outPath) {
  console.error("usage: node fetch-hpms-state.mjs <STATE_FIPS: 17|18|26|27|39|55> <out.ndjson>");
  process.exit(2);
}

class TooBigError extends Error {}
async function getJson(url, tries = 5) {
  for (let i = 0; i < tries; i++) {
    try {
      const r = await fetch(url, { signal: AbortSignal.timeout(150000) });
      if (!r.ok) throw new Error("HTTP " + r.status);
      const j = await r.json();
      if (j.error) {
        // "Unable to perform query" = the server gave up on a cell that
        // covers too many rows (measured live: ~55 s then HTTP 400 body).
        // Not transient — signal the caller to split, don't retry.
        if (JSON.stringify(j.error).includes("Unable to perform query")) throw new TooBigError();
        throw new Error("service error " + JSON.stringify(j.error).slice(0, 200));
      }
      return j;
    } catch (e) {
      if (e instanceof TooBigError || i === tries - 1) throw e;
      await new Promise(res => setTimeout(res, 3000 * (i + 1)));
    }
  }
}

function cellUrl([x0, y0, x1, y1]) {
  return `${BASE}/query?where=STATE_ID%3D${fips}` +
    `&geometry=${x0.toFixed(5)},${y0.toFixed(5)},${x1.toFixed(5)},${y1.toFixed(5)}` +
    `&geometryType=esriGeometryEnvelope&inSR=4326&spatialRel=esriSpatialRelIntersects` +
    `&outFields=OBJECTID,F_SYSTEM&returnGeometry=true&outSR=4326&geometryPrecision=6` +
    `&resultRecordCount=${PAGE_CAP}&f=geojson`;
}

await mkdir(dirname(outPath), { recursive: true });
const out = createWriteStream(outPath);
const seen = new Set();
// Seed with a grid rather than the whole state box — cells much bigger
// than ~1 degree just cost a slow server-side give-up before splitting.
const queue = [];
{
  const [x0, y0, x1, y1] = box, STEP = 1.0;
  for (let x = x0; x < x1; x += STEP)
    for (let y = y0; y < y1; y += STEP)
      queue.push([x, y, Math.min(x + STEP, x1), Math.min(y + STEP, y1)]);
}
let written = 0, cellsDone = 0, splits = 0;

async function worker() {
  for (;;) {
    const cell = queue.shift();
    if (!cell) {
      // queue may refill while other workers are mid-cell — linger briefly
      await new Promise(res => setTimeout(res, 500));
      if (!queue.length && inFlight === 0) return;
      continue;
    }
    inFlight++;
    try {
      const [x0, y0, x1, y1] = cell;
      const split = () => {
        const mx = (x0 + x1) / 2, my = (y0 + y1) / 2;
        queue.push([x0, y0, mx, my], [mx, y0, x1, my], [x0, my, mx, y1], [mx, my, x1, y1]);
        splits++;
      };
      let j;
      try { j = await getJson(cellUrl(cell)); }
      catch (e) {
        if (e instanceof TooBigError && (x1 - x0) > MIN_CELL_DEG) { split(); continue; }
        throw e;
      }
      const feats = j.features || [];
      const exceeded = !!((j.properties && j.properties.exceededTransferLimit) || j.exceededTransferLimit);
      if ((exceeded || feats.length >= PAGE_CAP) && (x1 - x0) > MIN_CELL_DEG) {
        split();
      } else {
        let lines = "";
        for (const f of feats) {
          const oid = f.properties && f.properties.OBJECTID;
          const cls = Math.trunc(Number(f.properties && f.properties.F_SYSTEM));
          if (oid == null || seen.has(oid) || !(cls >= 1 && cls <= 7) || !f.geometry) continue;
          seen.add(oid);
          lines += JSON.stringify({
            type: "Feature",
            tippecanoe: { minzoom: MINZOOM_BY_CLASS[cls] },
            properties: { F: cls },
            geometry: f.geometry,
          }) + "\n";
          written++;
        }
        if (lines) await new Promise((res, rej) => out.write(lines, e => e ? rej(e) : res()));
        cellsDone++;
        if (cellsDone % 25 === 0) console.log(`  ${cellsDone} cells done, ${splits} splits, ${written} features, queue ${queue.length}`);
      }
    } finally { inFlight--; }
  }
}
let inFlight = 0;

console.log(`FIPS ${fips}: quadtree harvest of ${JSON.stringify(box)}`);
await Promise.all(Array.from({ length: CONCURRENCY }, worker));
await new Promise(res => out.end(res));
console.log(`DONE: ${written} features, ${cellsDone} leaf cells, ${splits} splits -> ${outPath}`);
