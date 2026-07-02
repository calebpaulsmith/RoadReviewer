// Verifies web/index.html's rr-core classification logic (the JS port of
// modClassify.bas) against the live MDOT / INDOT / WisDOT / NTAD / TIGER
// services, using the confirmed test coordinates from CLAUDE.md §4.2,
// §4.2a and §4.2b. Runs the exact <script id="rr-core"> block shipped in
// the page — not a copy — so a passing run vouches for the committed file.
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
  "({ classifyPoint, detectState, parseCoordinates, federalAidVerdict, wisconsinLocalCategoryToFhwa })", ctx);

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

console.log("detectState:");
for (const [lat, lon, want] of [
  [42.28536, -85.57025, "MI"], [44.27, -83.52, "MI"], [46.5, -87.4, "MI"],   // Kalamazoo, Iosco, Marquette (UP)
  [39.7684, -86.1581, "IN"], [43.0389, -87.9065, "WI"], [45.169879, -89.102452, "WI"],
  [46.65, -90.86, "WI"], [36.16, -86.78, null],                               // Washburn Co, Nashville TN
]) check(`${lat},${lon} -> ${want}`, core.detectState(lat, lon) === want, "got " + core.detectState(lat, lon));

console.log("federalAidVerdict edge cases:");
check("rural minor collector", core.federalAidVerdict([6], false, 200) === "Non-federal aid - Rural Minor Collector");
check("urban minor collector", core.federalAidVerdict([6], true, 200) === "Federal aid - Urban Minor Collector");
check("mixed 7+5 rural", core.federalAidVerdict([7, 5], false, 200) === "Federal aid - Rural Major Collector");
check("code 0 review", core.federalAidVerdict([0], true, 200) === "Review - non-certified class, check manually");
check("WI 96 -> 5", core.wisconsinLocalCategoryToFhwa(96) === 5);

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
  ["WI", 45.4711, -89.7345, { verdictStarts: "Federal aid - Rural", classIncludes: "Major Collector" }],
  ["WI", 45.169879, -89.102452, { verdict: "Non-federal aid - Rural Minor Collector", urbanRural: "Rural" }],   // §7a fix #2 regression
  ["WI", 44.764850, -91.406533, { verdict: "Federal aid - Urban Interstate", acubName: "Eau Claire, WI" }],      // §7a fix #3 regression
  ["MN", 44.9778, -93.2650, { verdict: "ACUB only - class lookup not wired for this state", urbanRural: "Urban" }], // Minneapolis: state gate
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
  check(`${state} ${lat},${lon} -> ${r.verdict}` + (r.roadName ? ` [${r.roadName}]` : "") + (r.streets ? ` {${r.streets}}` : ""),
    problems.length === 0, problems.join("; "));
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
