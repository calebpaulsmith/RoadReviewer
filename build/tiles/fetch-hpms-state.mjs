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
// Each output feature carries properties.F — FHWA functional class 1-7
// (0/null dropped). The progressive class-by-zoom banding is applied by
// build-state-tiles.sh via a tippecanoe -j $zoom filter, NOT via the
// per-feature "tippecanoe":{"minzoom"} key: that key silently RATE-DROPS
// line features even at maxzoom (repro'd v2.49: 5 clean parallel lines ->
// 1 survivor, "dropped_by_rate" in tile strategies, -r1 doesn't help).

import { createWriteStream } from "node:fs";
import { mkdir } from "node:fs/promises";
import { dirname } from "node:path";

const BASE = "https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/HPMS_National_Current/FeatureServer/0";
const PAGE_CAP = 2000;        // the layer's maxRecordCount
const CONCURRENCY = 4;
const MIN_CELL_DEG = 0.005;   // never split below ~500 m — data-error guard

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
  // ROUTE_ID / mileposts / route name+number are the STATE's own LRS keys
  // (HPMS is built from the states' submissions) — baked into the tiles so
  // cached verdicts and exports can point back at the state inventory.
  return `${BASE}/query?where=STATE_ID%3D${fips}` +
    `&geometry=${x0.toFixed(5)},${y0.toFixed(5)},${x1.toFixed(5)},${y1.toFixed(5)}` +
    `&geometryType=esriGeometryEnvelope&inSR=4326&spatialRel=esriSpatialRelIntersects` +
    `&outFields=OBJECTID,F_SYSTEM,ROUTE_ID,BEGIN_POINT,END_POINT,RouteName,RouteNumber` +
    // NO resultRecordCount: the service stopped accepting it (republish,
    // 2026-09-14 evening) — the layer's own maxRecordCount (2000) caps the
    // page and exceededTransferLimit still signals the split.
    `&returnGeometry=true&outSR=4326&geometryPrecision=6&f=geojson`;
}

await mkdir(dirname(outPath), { recursive: true });
const out = createWriteStream(outPath);
const seen = new Set();
// Seed with a grid rather than the whole state box. Cell size matters a
// lot: dense-metro cells at ~1 degree don't answer at all (55 s server
// give-up each before we learn to split), while ~0.2-0.5 degree cells
// answer in seconds — capped + exceededTransferLimit — and split cheaply.
const queue = [];
{
  const [x0, y0, x1, y1] = box, STEP = 0.5;
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
        if (e instanceof TooBigError && (x1 - x0) > MIN_CELL_DEG) {
          console.log(`  give-up cell ${(x1 - x0).toFixed(2)}deg @ ${x0.toFixed(2)},${y0.toFixed(2)} -> split`);
          split(); continue;
        }
        throw e;
      }
      const feats = j.features || [];
      const exceeded = !!((j.properties && j.properties.exceededTransferLimit) || j.exceededTransferLimit);
      if ((exceeded || feats.length >= PAGE_CAP) && (x1 - x0) > MIN_CELL_DEG) {
        split();
      } else {
        let lines = "";
        for (const f of feats) {
          const p = f.properties || {};
          const oid = p.OBJECTID;
          const cls = Math.trunc(Number(p.F_SYSTEM));
          if (oid == null || seen.has(oid) || !(cls >= 1 && cls <= 7) || !f.geometry) continue;
          seen.add(oid);
          // Compact keys; null/empty attrs omitted so locals without a
          // route name cost nothing. R = state LRS ROUTE_ID; B/E =
          // begin/end mileposts; N = RouteName; RN = RouteNumber.
          const props = { F: cls };
          if (p.ROUTE_ID != null && String(p.ROUTE_ID).trim() !== "") props.R = String(p.ROUTE_ID).trim();
          if (p.BEGIN_POINT != null) props.B = Math.round(p.BEGIN_POINT * 1000) / 1000;
          if (p.END_POINT != null) props.E = Math.round(p.END_POINT * 1000) / 1000;
          if (p.RouteName != null && String(p.RouteName).trim() !== "") props.N = String(p.RouteName).trim();
          if (p.RouteNumber != null && p.RouteNumber !== 0) props.RN = p.RouteNumber;
          lines += JSON.stringify({ type: "Feature", properties: props, geometry: f.geometry }) + "\n";
          written++;
        }
        if (lines) await new Promise((res, rej) => out.write(lines, e => e ? rej(e) : res()));
        cellsDone++;
        if (cellsDone % 10 === 0) console.log(`  ${cellsDone} cells done, ${splits} splits, ${written} features, queue ${queue.length}`);
      }
    } finally { inFlight--; }
  }
}
let inFlight = 0;

console.log(`FIPS ${fips}: quadtree harvest of ${JSON.stringify(box)}`);
await Promise.all(Array.from({ length: CONCURRENCY }, worker));
await new Promise(res => out.end(res));
console.log(`DONE: ${written} features, ${cellsDone} leaf cells, ${splits} splits -> ${outPath}`);
