// How often does FHWA HPMS disagree with the state DOT's own class layer?
//
//   node build/audit-hpms-vs-state.mjs <MI|IN|WI|MN|IL|OH|all> [--n 150] [--buffer 50] [--out audit.csv]
//
// The web tool answers from HPMS tiles by default and Excel answers from the
// state layer, so every place the two differ is a place the two products can
// give different verdicts (WI STH 52, 2026-09-30). This samples real road
// segments out of HPMS, asks BOTH sources about the segment's midpoint, and
// compares them with the page's own code: rr-core is executed straight out of
// web/index.html (as verify-web-core.mjs does), so the state queries, the
// non-inventory twin drop and crossCheck are the shipped ones, not copies.
//
// Sampling: random 0.05-degree cells inside the state's box; from each cell a
// few random HPMS segments, half the sample from classes 1-6 and half locals
// (locals are ~85% of the network but the federal-aid line runs through the
// collectors, so a plain random sample would mostly compare Local with Local).
// Urban vs rural comes from the sampled segment's HPMS URBAN_ID (99999 =
// rural) rather than a live boundary query per point - it is applied to both
// sources alike, so it cannot manufacture a disagreement.
//
// Output: one CSV row per sampled point + a per-state summary. "disagree" =
// the two sources give different federal-aid answers; "class-differs" = same
// answer, different class; "none" = the state layer has no road within the
// buffer; "unclear" = the state layer itself is tied or non-certified there.
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import vm from "node:vm";

const here = dirname(fileURLToPath(import.meta.url));
const html = readFileSync(join(here, "..", "web", "index.html"), "utf8");
const m = html.match(/<script id="rr-core">([\s\S]*?)<\/script>/);
if (!m) { console.error("rr-core block not found in web/index.html"); process.exit(1); }

const UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36";
async function httpGetJson(url, tries = 3) {
  for (let i = 0; ; i++) {
    try {
      const r = await fetch(url, { headers: { "User-Agent": UA }, signal: AbortSignal.timeout(60000) });
      if (!r.ok) throw new Error("HTTP " + r.status);
      return await r.json();
    } catch (e) {
      if (i >= tries - 1) throw e;
      await new Promise(res => setTimeout(res, 1500 * (i + 1)));
    }
  }
}
const ctx = vm.createContext({ httpGetJson, console });
ctx.RR_DOUBLE_CHECK = true;   // have the state queries return route ids, as the page does with the double-check on
vm.runInContext(m[1], ctx, { filename: "rr-core (from web/index.html)" });
const core = vm.runInContext("({ NFC_WIRED, crossCheck, dropNonInventoryTwins, runQuery, featureEntries, cleanStr, functionalSystemLabel })", ctx);

const HPMS = "https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/HPMS_National_Current/FeatureServer/0";
const STATES = {
  MI: { fips: 26, box: [-90.5, 41.6, -82.0, 48.4] }, WI: { fips: 55, box: [-93.0, 42.4, -86.1, 47.4] },
  MN: { fips: 27, box: [-97.3, 43.4, -89.4, 49.5] }, IL: { fips: 17, box: [-91.6, 36.9, -87.0, 42.6] },
  IN: { fips: 18, box: [-88.2, 37.7, -84.7, 41.8] }, OH: { fips: 39, box: [-84.9, 38.3, -80.4, 42.0] },
};

const args = process.argv.slice(2);
const opt = (name, def) => { const i = args.indexOf("--" + name); return i >= 0 ? args[i + 1] : def; };
const which = (args[0] || "").toUpperCase();
const states = which === "ALL" ? Object.keys(STATES) : STATES[which] ? [which] : null;
if (!states) { console.error("usage: node build/audit-hpms-vs-state.mjs <MI|IN|WI|MN|IL|OH|all> [--n 150] [--buffer 50] [--out audit.csv]"); process.exit(2); }
const N = parseInt(opt("n", "150"), 10), BUFFER = parseInt(opt("buffer", "50"), 10);
const outPath = opt("out", "hpms-vs-state-audit.csv");
const CELL = 0.05, PER_CELL = 3, WORKERS = 3;

// Random HPMS segments of one class group ("F_SYSTEM<7" or "F_SYSTEM=7").
async function sampleCell(st, classWhere) {
  const { fips, box } = STATES[st];
  const x = box[0] + Math.random() * (box[2] - box[0] - CELL), y = box[1] + Math.random() * (box[3] - box[1] - CELL);
  const url = `${HPMS}/query?where=${encodeURIComponent(`STATE_ID=${fips} AND ${classWhere}`)}` +
    `&geometry=${x.toFixed(4)},${y.toFixed(4)},${(x + CELL).toFixed(4)},${(y + CELL).toFixed(4)}` +
    `&geometryType=esriGeometryEnvelope&inSR=4326&spatialRel=esriSpatialRelIntersects` +
    `&outFields=OBJECTID,F_SYSTEM,URBAN_ID,FACILITY_TYPE,RouteName&returnGeometry=true&outSR=4326&geometryPrecision=6&f=json`;
  let j;
  try { j = await httpGetJson(url, 1); } catch { return []; }
  const feats = (j.features || []).filter(f => f.geometry && f.geometry.paths && f.geometry.paths[0] && f.geometry.paths[0].length);
  const picks = [];
  for (let i = 0; i < PER_CELL && feats.length; i++) {
    const f = feats.splice(Math.floor(Math.random() * feats.length), 1)[0];
    const path = f.geometry.paths[0], mid = path[path.length >> 1];
    picks.push({ lat: +mid[1].toFixed(6), lon: +mid[0].toFixed(6), urban: f.attributes.URBAN_ID != null && f.attributes.URBAN_ID < 99999,
      sampledClass: f.attributes.F_SYSTEM, routeName: core.cleanStr(f.attributes.RouteName) });
  }
  return picks;
}

async function hpmsSegments(lat, lon) {
  const j = await core.runQuery(HPMS, lat, lon, "F_SYSTEM,FACILITY_TYPE,RouteNumber,RouteName,ROUTE_ID", "1=1", BUFFER, true);
  const segs = core.featureEntries(j, lat, lon)
    .filter(e => Number.isFinite(Number(e.attrs.F_SYSTEM)) && e.attrs.F_SYSTEM !== null)
    .map(e => ({ code: Math.trunc(Number(e.attrs.F_SYSTEM)), distFt: e.distFt, name: core.cleanStr(e.attrs.RouteName),
      routeNumber: e.attrs.RouteNumber ? String(e.attrs.RouteNumber) : "", routeId: core.cleanStr(e.attrs.ROUTE_ID), facility: Math.trunc(Number(e.attrs.FACILITY_TYPE)) || 0 }));
  return core.dropNonInventoryTwins(segs).sort((a, b) => a.distFt - b.distFt);
}

const rows = [["state", "lat", "lon", "urban", "status", "hpms_class", "state_class", "hpms_route", "note", "matched_by"]];
const by = {};
const q = v => '"' + String(v ?? "").replaceAll('"', '""') + '"';
const summary = {};

for (const st of states) {
  const tally = { sampled: 0, agree: 0, "class-differs": 0, disagree: 0, none: 0, unclear: 0, "hpms-unclear": 0, error: 0 };
  summary[st] = tally;
  // ---- collect the sample ----
  const points = [];
  for (const [classWhere, want] of [["F_SYSTEM<7", Math.ceil(N / 2)], ["F_SYSTEM=7", Math.floor(N / 2)]]) {
    const got = [];
    for (let tries = 0; got.length < want && tries < want * 6; tries += WORKERS) {
      const batches = await Promise.all(Array.from({ length: WORKERS }, () => sampleCell(st, classWhere)));
      for (const b of batches) for (const p of b) if (got.length < want) got.push(p);
    }
    points.push(...got);
  }
  console.log(`${st}: ${points.length} sample points (buffer ${BUFFER} ft)`);

  // ---- compare ----
  let next = 0;
  await Promise.all(Array.from({ length: WORKERS }, async () => {
    for (;;) {
      const p = points[next++];
      if (!p) return;
      tally.sampled++;
      let status, hp = "", sc = "", note = "", matchedBy = "";
      try {
        const [h, s] = await Promise.all([hpmsSegments(p.lat, p.lon), core.NFC_WIRED[st](p.lat, p.lon, BUFFER)]);
        // HPMS itself has roads of both federal-aid answers within 30 ft of
        // its closest one here (an intersection, or a twin that could not be
        // dropped): the page already calls that a review, so it says nothing
        // about whether the two SOURCES differ.
        const fed = c => c >= 1 && c <= 5 || (c === 6 && p.urban);
        if (!h.length || h.some(s => s.code !== 0 && s.distFt - h[0].distFt < 30 && fed(s.code) !== fed(h[0].code))) { status = "hpms-unclear"; }
        else {
          const cc = core.crossCheck(h, s.segments, p.urban);
          status = cc.status === "other-only" ? "agree" : cc.status;
          matchedBy = cc.by || "";
          if (cc.by) by[cc.by] = (by[cc.by] || 0) + 1;
          hp = core.functionalSystemLabel(h[0].code);
          if (cc.code != null) sc = core.functionalSystemLabel(cc.code);
        }
      } catch (e) { status = "error"; note = e.message; }
      tally[status] = (tally[status] || 0) + 1;
      rows.push([st, p.lat, p.lon, p.urban ? "urban" : "rural", status, q(hp), q(sc), q(p.routeName), q(note), matchedBy]);
      if (tally.sampled % 25 === 0) console.log(`  ${st}: ${tally.sampled}/${points.length} compared, ${tally.disagree} disagree`);
    }
  }));
}

writeFileSync(outPath, rows.map(r => r.join(",")).join("\n") + "\n");
console.log("\nstate  sampled  agree  class-differs  DISAGREE  state-has-no-road  state-unclear  hpms-unclear  errors");
for (const [st, t] of Object.entries(summary)) {
  const compared = t.agree + t["class-differs"] + t.disagree;
  console.log(`${st.padEnd(5)}  ${String(t.sampled).padStart(7)}  ${String(t.agree).padStart(5)}  ${String(t["class-differs"]).padStart(13)}  ` +
    `${String(t.disagree).padStart(8)}  ${String(t.none).padStart(17)}  ${String(t.unclear).padStart(13)}  ${String(t["hpms-unclear"]).padStart(12)}  ${String(t.error).padStart(6)}` +
    (compared ? `   -> ${(100 * t.disagree / compared).toFixed(1)}% of compared points disagree on federal aid` : ""));
}
console.log(`\nmatched by route id: ${by.id || 0}   by distance (no shared id): ${by.distance || 0}`);
console.log(`rows -> ${outPath}`);
