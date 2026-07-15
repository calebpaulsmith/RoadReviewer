// Verifies web/index.html's rr-core classification logic (the JS port of
// modClassify.bas) against the live MDOT / INDOT / WisDOT / MnDOT / IDOT /
// ODOT / NTAD / TIGER services, using the confirmed test coordinates from
// CLAUDE.md §4.2, §4.2a, §4.2b and §4.2c-e. Runs the exact
// <script id="rr-core"> block shipped in the page — not a copy — so a
// passing run vouches for the committed file.
//
//   node build/verify-web-core.mjs
//
// Network goes through curl so the sandbox HTTPS proxy + CA bundle are
// honored without any Node fetch/agent configuration.

import { readFileSync } from "node:fs";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const execFileP = promisify(execFile);
const here = dirname(fileURLToPath(import.meta.url));
const html = readFileSync(join(here, "..", "web", "index.html"), "utf8");

const m = html.match(/<script id="rr-core">([\s\S]*?)<\/script>/);
if (!m) { console.error("FAIL: <script id=\"rr-core\"> not found in web/index.html"); process.exit(1); }

async function httpGetJson(url) {
  const { stdout } = await execFileP("curl", ["-sS", "--max-time", "45", url], { maxBuffer: 64 * 1024 * 1024 });
  return JSON.parse(stdout);
}

const ctx = vm.createContext({ httpGetJson, console });
vm.runInContext(m[1], ctx, { filename: "rr-core (from web/index.html)" });
const core = vm.runInContext(
  "({ classifyPoint, detectState, parseCoordinates, computeVerdict, classIsFederal, mergeRoadList, minDistanceFt, wisconsinLocalCategoryToFhwa })", ctx);

let pass = 0, fail = 0;
function check(label, ok, detail) {
  if (ok) { pass++; console.log("  ok   " + label); }
  else { fail++; console.log("  FAIL " + label + (detail ? " — " + detail : "")); }
}

/* ---------- offline unit checks ---------- */
console.log("parseCoordinates:");
{
  const pts = core.parseCoordinates(
    "42.28536, -85.57025\nCulvert on Q Ave\t42.6911\t-84.5360\nSite 12, 44.2700, -83.5200\n-85.5 42.3 name after\ngarbage line\n");
  check("parses 4 valid + 1 invalid", pts.filter(p => !p.invalid).length === 4 && pts.filter(p => p.invalid).length === 1,
    JSON.stringify(pts));
  check("bare pair", pts[0].lat === 42.28536 && pts[0].lon === -85.57025);
  check("tab + name-before", pts[1].name === "Culvert on Q Ave" && pts[1].lat === 42.6911);
  check("leading site number not mistaken for lat", pts[2].lat === 44.27 && pts[2].lon === -83.52);
  check("swapped lon,lat accepted", pts[3].lat === 42.3 && pts[3].lon === -85.5);
}

console.log("detectState (all six states, PR #36):");
for (const [lat, lon, want] of [
  [42.28536, -85.57025, "MI"], [44.27, -83.52, "MI"], [46.5, -87.4, "MI"],   // Kalamazoo, Iosco, Marquette (UP)
  [39.7684, -86.1581, "IN"], [43.0389, -87.9065, "WI"], [45.169879, -89.102452, "WI"],
  [46.65, -90.86, "WI"], [36.16, -86.78, null],                               // Washburn Co, Nashville TN
  [44.9778, -93.2650, "MN"], [45.822764, -95.222414, "MN"],                   // Minneapolis, Douglas Co
  [41.8781, -87.6298, "IL"], [40.12, -87.63, "IL"],                           // Chicago, Danville (east border)
  [39.9612, -82.9988, "OH"], [41.87, -80.80, "OH"],                           // Columbus, Ashtabula (lakeshore)
  [41.92, -83.40, "MI"],                                                      // Monroe MI (not OH's box)
]) check(`${lat},${lon} -> ${want}`, core.detectState(lat, lon) === want, "got " + core.detectState(lat, lon));

console.log("computeVerdict (closest-road + ambiguity model, port of modClassify.ComputeVerdict):");
{
  const seg = (code, distFt, name = "") => ({ code, distFt, name });
  const v = (segs, urban, boundary) => core.computeVerdict(segs, urban, boundary, 250);
  check("closest rural local -> green", v([seg(7, 5)], false, false).verdict === "Non-federal aid - Rural Local");
  check("closest urban local -> green", v([seg(7, 5)], true, false).verdict === "Non-federal aid - Urban Local");
  check("closest urban minor collector -> red", v([seg(6, 5)], true, false).verdict === "Federal aid - Urban Minor Collector");
  check("closest rural minor collector -> green (correct label)", v([seg(6, 5)], false, false).verdict === "Non-federal aid - Rural Minor Collector");
  check("red stays red with local nearby", v([seg(5, 5), seg(7, 10)], false, false).verdict === "Federal aid - Rural Major Collector");
  check("interstate gets urban/rural prefix", v([seg(1, 5)], true, false).verdict === "Federal aid - Urban Interstate");
  check("local closest + federal within 30 ft -> Second road close",
    v([seg(7, 5), seg(5, 20)], false, false).verdict === "Review - Second road close");
  check("local closest + federal beyond 30 ft -> Nearby FHWA road",
    v([seg(7, 5), seg(5, 100)], false, false).verdict === "Review - Nearby FHWA road");
  check("second road within 30 ft but NOT federal stays green",
    v([seg(7, 5), seg(6, 20)], false, false).verdict === "Non-federal aid - Rural Local");
  check("second road within 30 ft, urban 6 IS federal -> yellow",
    v([seg(7, 5), seg(6, 20)], true, false).reason === "Second road close");
  check("minor collector on urban boundary edge -> yellow",
    v([seg(6, 5)], false, true).verdict === "Review - Urban boundary edge");
  check("local on boundary edge stays green (only class 6 flips)",
    v([seg(7, 5)], false, true).verdict === "Non-federal aid - Rural Local");
  check("non-certified closest -> review", v([seg(0, 5)], true, false).verdict === "Review - non-certified class, check manually");
  check("no segments -> review", v([], false, false).verdict === "Review - no road within 250 ft" && v([], false, false).reason === "No road found");
  check("WI 96 -> 5", core.wisconsinLocalCategoryToFhwa(96) === 5);
}

console.log("distance helpers:");
{
  // A straight N-S segment 0.00125 deg east of the point at lat ~42.3:
  // 0.00125 deg lon * 111320 m/deg * cos(42.28536deg) = 102.94 m = 337.7 ft.
  const d = core.minDistanceFt({ paths: [[[-85.569, 42.28], [-85.569, 42.29]]] }, 42.28536, -85.57025);
  check("point-to-polyline distance ~338 ft (equirectangular)", Math.abs(d - 337.7) < 2, "got " + d.toFixed(1));
  check("no geometry -> Infinity", core.minDistanceFt(undefined, 42, -85) === Infinity);
  const merged = core.mergeRoadList([{ name: "Main St", distFt: 50 }, { name: "MAIN ST", distFt: 10 }, { name: "Q Ave", distFt: 30 }]);
  check("road list dedups by name keeping nearest, sorted",
    merged.length === 2 && merged[0].name === "MAIN ST" && merged[0].distFt === 10 && merged[1].name === "Q Ave");
}

/* ---------- live service checks (§4.2 / §4.2a / §4.2b test coordinates) ---------- */
const CASES = [
  // [state, lat, lon, assertions...]
  ["MI", 42.28536, -85.57025, { verdict: "Federal aid - Urban Minor Collector", urbanRural: "Urban", acubName: "Kalamazoo, MI", classIncludes: "Minor Collector" }],
  ["MI", 42.6911, -84.5360, { verdict: "Non-federal aid - Urban Local", urbanRural: "Urban", acubName: "Lansing, MI" }],
  ["MI", 44.2700, -83.5200, { verdict: "Non-federal aid - Rural Local", urbanRural: "Rural" }],
  ["IN", 39.7684, -86.1581, { verdictStarts: "Federal aid -", urbanRural: "Urban", classIncludes: "Minor Collector" }],
  ["IN", 39.4234, -86.7628, { verdictStarts: "Federal aid -", classIncludes: "Other Principal Arterial" }],
  ["IN", 39.9876, -86.0128, { verdictStarts: "Non-federal aid -", classIncludes: "Local" }],
  ["WI", 43.0389, -87.9065, { verdict: "Federal aid - Urban Minor Arterial", urbanRural: "Urban", acubName: "Milwaukee, WI" }],
  // Local-first WI order (PR "WI layer swap"): the point sits on W Wisconsin Ave
  // (Rural Minor Collector, ~26 ft) — much closer than STH 86 (Major Collector,
  // ~199 ft), which the stub-triggered trunk fallback still surfaces. Closest
  // road drives the base verdict (non-federal), STH 86 downgrades it to yellow.
  ["WI", 45.4711, -89.7345, { verdict: "Review - Nearby FHWA road", urbanRural: "Rural", classIncludes: "Major Collector" }],
  ["WI", 45.169879, -89.102452, { verdict: "Non-federal aid - Rural Minor Collector", urbanRural: "Rural" }],   // §7a fix #2 regression
  ["WI", 44.764850, -91.406533, { verdict: "Federal aid - Urban Interstate", acubName: "Eau Claire, WI" }],      // §7a fix #3 regression
  // MN/IL/OH (PR #36, §4.2c-e test coordinates - all live-verified 2026-07-15):
  ["MN", 44.9531, -93.1668, { verdict: "Federal aid - Urban Minor Arterial", urbanRural: "Urban", acubName: "Minneapolis--St. Paul, MN", classIncludes: "Minor Arterial" }],  // Snelling Ave, St Paul
  ["MN", 44.9260, -93.2570, { verdict: "Non-federal aid - Urban Local", urbanRural: "Urban" }],                  // Minneapolis residential
  ["MN", 45.822764, -95.222414, { verdict: "Non-federal aid - Rural Local", urbanRural: "Rural" }],              // rural Douglas County
  ["IL", 41.9020, -87.6870, { verdict: "Federal aid - Urban Other Principal Arterial", urbanRural: "Urban", acubName: "Chicago, IL--IN" }],  // Western Ave, Chicago
  ["IL", 41.9430, -87.7010, { verdict: "Non-federal aid - Urban Local", urbanRural: "Urban" }],                  // Chicago residential
  ["IL", 40.165157, -89.434236, { verdict: "Non-federal aid - Rural Local", urbanRural: "Rural" }],              // rural Logan County
  ["OH", 40.0150, -82.9990, { verdict: "Federal aid - Urban Other Principal Arterial", urbanRural: "Urban", acubName: "Columbus, OH" }],     // N High St / US 23, Columbus
  ["OH", 40.0855, -83.0170, { verdict: "Non-federal aid - Urban Local", urbanRural: "Urban" }],                  // Columbus residential
  ["OH", 40.320352, -83.302785, { verdict: "Non-federal aid - Rural Local", urbanRural: "Rural" }],              // rural Marion County
  ["TN", 36.16, -86.78, { verdict: "ACUB only - class lookup not wired for this state", urbanRural: "Urban" }],  // state gate (unwired state)
];

console.log("live classification (" + CASES.length + " points):");
for (const [state, lat, lon, want] of CASES) {
  let r;
  try { r = await core.classifyPoint(lat, lon, state); }
  catch (e) { check(`${state} ${lat},${lon}`, false, "threw " + e.message); continue; }
  const problems = [];
  if (want.verdict && r.verdict !== want.verdict) problems.push(`verdict "${r.verdict}" != "${want.verdict}"`);
  if (want.verdictStarts && !r.verdict.startsWith(want.verdictStarts)) problems.push(`verdict "${r.verdict}" !^ "${want.verdictStarts}"`);
  if (want.urbanRural && r.urbanRural !== want.urbanRural) problems.push(`urbanRural "${r.urbanRural}" != "${want.urbanRural}"`);
  if (want.acubName && r.acubName !== want.acubName) problems.push(`acubName "${r.acubName}" != "${want.acubName}"`);
  if (want.classIncludes && !r.classLabel.includes(want.classIncludes)) problems.push(`classLabel "${r.classLabel}" !~ "${want.classIncludes}"`);
  // PR #24 parity: wired-state results must carry real per-segment distances
  // and a nearest-first merged road list with "(N ft)" suffixes.
  if (["MI", "IN", "WI", "MN", "IL", "OH"].includes(state)) {
    if (!r.segments.length || !r.segments.every(s => Number.isFinite(s.distFt))) problems.push("segments missing finite distFt");
    if (r.roadName && !/\(\d+ ft\)/.test(r.roadName)) problems.push(`roadName "${r.roadName}" missing (N ft) distances`);
    const dists = r.roadList.map(x => x.distFt);
    if (dists.some((d, i) => i > 0 && d < dists[i - 1])) problems.push("roadList not sorted nearest-first");
  }
  check(`${state} ${lat},${lon} -> ${r.verdict}` + (r.reviewReason ? ` (${r.reviewReason})` : "") +
    (r.roadName ? ` [${r.roadName}]` : "") + (r.streets ? ` {${r.streets}}` : ""),
    problems.length === 0, problems.join("; "));
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
