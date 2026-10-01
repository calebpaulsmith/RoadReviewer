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
  "({ classifyPoint, detectState, parseCoordinates, computeVerdict, classIsFederal, mergeRoadList, minDistanceFt, wisconsinLocalCategoryToFhwa, " +
  "crossCheck, applyCrossCheck, dropNonInventoryTwins, parseAddressLine, parseRoadLine, streetMatchScore, streetFromTiger, streetFromGeocoder, closestPointOnPaths, pointAlongPath, pathLengthM, sanitizeGeocodeResponse })", ctx);

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

/* DMS / degrees-decimal-minutes. Every line below is the §4.2 Kalamazoo
   federal-aid test point written a different way. The signed forms are
   regressions: before DMS was parsed, "42 17 07.3, -85 34 12.9" read as the
   last in-range NUMBER pair — 34, -85 — and classified a point in Alabama
   without a word of warning. The name cases guard the hemisphere letters,
   which also live inside ordinary words ("Ave", "Rowe", "Escanaba"). */
console.log("parseCoordinates — degrees/minutes/seconds:");
{
  const K = (p, label, name) => {
    const okC = p && !p.invalid && Math.abs(p.lat - 42.28536) < 0.002 && Math.abs(p.lon + 85.57025) < 0.002;
    check(label, okC && (name === undefined || p.name === name),
      p ? `${p.invalid ? "INVALID" : p.lat + "," + p.lon} name=${JSON.stringify(p && p.name)}` : "no point");
  };
  const one = t => core.parseCoordinates(t)[0];
  K(one(`42°17'07.3"N 85°34'12.9"W`), "symbols + NSEW");
  K(one(`42°17'07.3"N, 85°34'12.9"W`), "symbols, comma-separated");
  K(one(`N42°17'07.3" W85°34'12.9"`), "hemisphere first");
  K(one("42 17 07.3 N, 85 34 12.9 W"), "spaces + NSEW, no symbols");
  K(one("42 17 07.3, -85 34 12.9"), "spaces + signed degrees (was 34,-85)");
  K(one(`42° 17.122' N, 85° 34.215' W`), "degrees + decimal minutes");
  K(one("42 17.122, -85 34.215"), "decimal minutes, signed (was 34.215,-85)");
  K(one(`Culvert on Q Ave 42°17'07.3"N 85°34'12.9"W`), "name before", "Culvert on Q Ave");
  K(one(`42°17'07.3"N 85°34'12.9"W, washout`), "name after", "washout");
  K(one("Rowe Rd 42 17 07.3, -85 34 12.9"), "'e' in a name is not EAST", "Rowe Rd");
  K(one("42 17 07.3 N, 85 34 12.9 W Escanaba culvert"), "trailing word keeps its first letter", "Escanaba culvert");
  // A site name ending in a direction word donates its last letter to the
  // coordinate as a hemisphere, which flips the latitude negative and used to
  // lose the whole line; the parser retries without a leading hemisphere.
  K(one(`Rose Drive W 42\u00b017'07.3"N 85\u00b034'12.9"W`), "name ending in a direction word", "Rose Drive W");
  K(one("Rose Drive S 42 17 07.3 N, 85 34 12.9 W"), "name ending in S", "Rose Drive S");
  K(one(`E Rose Dr 42\u00b017'07.3"N 85\u00b034'12.9"W`), "name starting with a direction word", "E Rose Dr");
  // The same names on DECIMAL lines (the common case) must be untouched.
  for (const [line, name] of [["42.28536, -85.57025 E Rose Dr", "E Rose Dr"],
                              ["Rose Drive W 42.28536, -85.57025", "Rose Drive W"],
                              ["Rose Dr S, 42.28536, -85.57025", "Rose Dr S"],
                              ["100 S Rose Dr 42.28536, -85.57025", "100 S Rose Dr"],
                              ["42.28536, -85.57025 US 131", "US 131"],
                              ["CR 42 N 42.28536, -85.57025", "CR 42 N"]])
    K(one(line), `decimal + direction name: ${JSON.stringify(line)}`, name);
  // Addresses are NOT accepted (deferred: the Census geocoder sends no CORS
  // header, so it can't be called from the page). They must stay INVALID
  // rather than have a street number and a ZIP read as a coordinate.
  for (const a of ["123 W Main St, Kalamazoo, MI 49007", "5201 Portage Rd, Portage MI 49002",
                   "8500 N 32nd St, Richland, MI", "2200 S 1700 W, Salt Lake City UT"])
    check(`address rejected: ${a}`, (one(a) || { invalid: true }).invalid === true);
}

/* Lines with no coordinate are sorted locally into address / road / unknown
   so the page can offer the right next step. Detection sends nothing. */
console.log("parseCoordinates — addresses and road names (detected locally):");
{
  const kind = l => (core.parseCoordinates(l)[0] || {}).kind;
  for (const [l, k] of [
    ["5201 Portage Rd, Portage, MI 49002", "address"], ["123 W Main St, Kalamazoo, MI 49007", "address"],
    ["8500 N 32nd St, Richland, MI", "address"], ["5201 Portage Rd Portage MI 49002", "address"],
    ["100 S Rose Dr", "address"], ["2200 S 1700 W, Salt Lake City UT", "address"],
    ["Portage Rd, Portage MI", "road"], ["M-43, Kalamazoo County", "road"], ["CR 550 N, Hamilton County IN", "road"],
    ["US 131, Kalamazoo", "road"], ["County Road 12, Iosco County MI", "road"],
    ["Q Ave", "road"], ["Portage Rd", "road"], ["garbage line", "unknown"], ["Kalamazoo culvert", "unknown"], ["100 Q", "unknown"],
  ]) check(`${k.padEnd(7)} <- ${JSON.stringify(l)}`, kind(l) === k, "got " + kind(l));
  const a = core.parseAddressLine("8500 N 32nd St, Richland, MI");
  check("address parts: house, preDir, base, type, city, state", a && a.house === "8500" && a.street.preDir === "N" && a.street.base === "32ND"
    && a.street.sufType === "ST" && a.city === "Richland" && a.state === "MI", JSON.stringify(a));
  const b = core.parseAddressLine("5201 Portage Rd Portage MI 49002");
  check("no-comma address splits at the street type", b && b.street.base === "PORTAGE" && b.city === "Portage" && b.zip === "49002", JSON.stringify(b));
  const r = core.parseRoadLine("CR 550 N, Hamilton County IN");
  check("route road line keeps the route and the place", r && r.street.route && r.street.base === "CR 550" && r.place === "Hamilton County" && r.state === "IN", JSON.stringify(r));
}

/* The geocoder's parsed street vs TIGER's street parts: the match that
   makes an address snap onto the street it names, not the nearest line. */
console.log("street matching + snapping geometry:");
{
  const want = core.streetFromGeocoder({ streetName: "PORTAGE", suffixType: "RD", preDirection: "", suffixDirection: "" });
  const sc = a => core.streetMatchScore(want, core.streetFromTiger(a));
  check("Portage Rd matches the geocoder's PORTAGE / RD", sc({ BASENAME: "Portage", SUFTYPEABRV: "Rd" }) >= 2);
  check("Airview Blvd 20 ft away does not", sc({ BASENAME: "Airview", SUFTYPEABRV: "Blvd" }) === 0);
  check("Portage St scores below Portage Rd", sc({ BASENAME: "Portage", SUFTYPEABRV: "St" }) < sc({ BASENAME: "Portage", SUFTYPEABRV: "Rd" }));
  const path = [[-85.5710, 42.2850], [-85.5700, 42.2850], [-85.5690, 42.2850]];
  const cp = core.closestPointOnPaths([path], 42.2852, -85.5701);
  check("closest point on a path lands on the line ~73 ft away", cp && Math.abs(cp.lat - 42.2850) < 1e-6 && cp.distFt > 70 && cp.distFt < 75, JSON.stringify(cp));
  const f = core.pointAlongPath(path, cp.segIdx, cp.t, 45.72), bk = core.pointAlongPath(path, cp.segIdx, cp.t, -45.72);
  check("±150 ft along the path stays on it, ~300 ft apart", Math.abs(f[0] - 42.2850) < 1e-6 && Math.abs(bk[0] - 42.2850) < 1e-6
    && Math.abs((f[1] - bk[1]) * 111320 * Math.cos(42.285 * Math.PI / 180) - 91.44) < 1);
  check("walking past the end clamps to the endpoint", core.pointAlongPath(path, 1, 0.9, 500)[1] === -85.5690);
  const good = { result: { addressMatches: [{ matchedAddress: "X", coordinates: { x: -85.56, y: 42.24 }, addressComponents: { streetName: "PORTAGE", suffixType: "RD" } }] } };
  const s = core.sanitizeGeocodeResponse(good);
  check("geocoder reply reduced to whitelisted fields only", s.length === 1 && Object.keys(s[0]).sort().join() === "components,lat,lon,matchedAddress,tigerLineId");
  check("junk geocoder reply throws", (() => { try { core.sanitizeGeocodeResponse({ nope: 1 }); return false; } catch { return true; } })());
  check("non-numeric / out-of-range coordinates dropped", core.sanitizeGeocodeResponse({ result: { addressMatches: [{ coordinates: { x: "a", y: 1 } }, { coordinates: { x: 999, y: 1 } }] } }).length === 0);
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
  // A TIE (same distance, within CLASS_TIE_FEET) with differing outcomes is a
  // review whichever segment came first — WI STH 52 in HPMS: the inventory
  // record is class 6, the non-inventory direction record class 3, both 25 ft.
  check("tie with different outcomes -> Conflicting classes (federal first)",
    v([seg(3, 25), seg(6, 25)], false, false).verdict === "Review - Conflicting classes");
  check("tie with different outcomes -> Conflicting classes (non-federal first)",
    v([seg(6, 25), seg(3, 25)], false, false).reason === "Conflicting classes");
  check("tie, both federal -> red", v([seg(1, 0), seg(2, 0)], false, false).verdict === "Federal aid - Rural Interstate");
  check("tie, both non-federal -> green", v([seg(7, 0), seg(6, 0)], false, false).verdict === "Non-federal aid - Rural Local");
  check("tie with a non-certified segment is not a conflict", v([seg(3, 0), seg(0, 0)], false, false).verdict === "Federal aid - Rural Other Principal Arterial");
  check("4 ft apart is not a tie", v([seg(3, 20), seg(7, 24)], false, false).verdict === "Federal aid - Rural Other Principal Arterial");
  // HPMS non-inventory twins (FACILITY_TYPE 6 beside its type-2 record on
  // the same centerline): the type-6 record is dropped, so STH 52 reads 6.
  {
    const tw = (code, distFt, facility, routeNumber, name) => ({ code, distFt, facility, routeNumber, name });
    const sth52 = core.dropNonInventoryTwins([tw(3, 25, 6, "52", "STH  052W"), tw(6, 25, 2, "52", "STH  052E")]);
    check("STH 52: type-6 twin dropped, inventory record left", sth52.length === 1 && sth52[0].code === 6);
    check("STH 52 after the twin drop -> Rural Minor Collector",
      v(sth52, false, false).verdict === "Non-federal aid - Rural Minor Collector");
    check("twin matched by name stem when there is no route number",
      core.dropNonInventoryTwins([tw(3, 25, 6, "", "STH  052W"), tw(6, 25, 2, "", "STH 052E")]).length === 1);
    check("twin matched by route id apart from the -D / -I suffix (MnDOT)",
      core.dropNonInventoryTwins([{ code: 5, distFt: 10, facility: 6, routeId: "0400006595000005-D" },
                                  { code: 6, distFt: 10, facility: 2, routeId: "0400006595000005-I" }]).map(s => s.code).join() === "6");
    check("type 6 with no twin is kept (WisDOT files most local streets as 6)",
      core.dropNonInventoryTwins([tw(7, 10, 6, "", "MAIN ST")]).length === 1);
    check("two type-6 records are both kept",
      core.dropNonInventoryTwins([tw(7, 10, 6, "5", ""), tw(6, 10, 6, "5", "")]).length === 2);
    check("a type-6 record on a different route is kept",
      core.dropNonInventoryTwins([tw(3, 25, 6, "52", "STH 052W"), tw(6, 25, 2, "64", "STH 064E")]).length === 2);
    check("a same-route record 40 ft away is not a twin (divided highway)",
      core.dropNonInventoryTwins([tw(3, 65, 6, "52", ""), tw(6, 25, 2, "52", "")]).length === 2);
  }
  // Double-check: the verdict's source against the other class source.
  {
    const cc = (p, o, urban = false) => core.crossCheck(p, o, urban);
    check("cross-check: same class -> agree", cc([seg(6, 25)], [seg(6, 24)]).status === "agree");
    check("cross-check: 6 vs 3 rural -> disagree", cc([seg(6, 25)], [seg(3, 25)]).status === "disagree");
    check("cross-check: 4 vs 3 -> class-differs (both federal)", cc([seg(4, 5)], [seg(3, 5)]).status === "class-differs");
    check("cross-check: 6 vs 7 rural -> class-differs (both non-federal)", cc([seg(6, 5)], [seg(7, 5)]).status === "class-differs");
    check("cross-check: 6 vs 7 urban -> disagree", cc([seg(6, 5)], [seg(7, 5)], true).status === "disagree");
    check("cross-check: other source finds no road -> none", cc([seg(6, 5)], []).status === "none");
    check("cross-check: other source tied, one of the two matches -> agree", cc([seg(6, 5)], [seg(3, 25), seg(6, 25)]).status === "agree");
    check("cross-check: intersection drawn 4 ft apart is not a disagreement (USH 14 at Autumn Dr)",
      cc([seg(3, 0), seg(7, 5)], [seg(7, 0), seg(3, 4)], true).status === "agree");
    check("cross-check: a matching road 40 ft past the other source's closest does not count",
      cc([seg(3, 0)], [seg(7, 0), seg(3, 40)]).status === "disagree");
    check("cross-check: near-closest match of a different class -> class-differs",
      cc([seg(3, 0)], [seg(7, 0), seg(4, 10)]).status === "class-differs");
    check("cross-check: other source non-certified -> unclear", cc([seg(6, 5)], [seg(0, 5)]).status === "unclear");
    check("cross-check: this source has no class -> other-only", cc([], [seg(7, 30)]).status === "other-only");
    check("cross-check: the closest segment speaks for each source", cc([seg(7, 5), seg(3, 90)], [seg(7, 6), seg(3, 80)]).status === "agree");
    // Route ids: the same road is compared, whatever each source draws closest.
    const sid = (code, distFt, routeId) => ({ code, distFt, routeId });
    const byId = cc([sid(3, 0, "42510"), sid(7, 20, "A")], [sid(7, 0, "4112437"), sid(3, 45, "42510")], true);
    check("cross-check by route id: same road agrees even when the other source draws another road 45 ft closer",
      byId.status === "agree" && byId.by === "id", JSON.stringify(byId));
    const idDis = cc([sid(7, 0, "CGRECR00012**C")], [sid(6, 2, "cgrecr00012**c"), sid(7, 3, "X")], true);
    check("cross-check by route id: same id, different federal-aid answer -> disagree (id match is case/space-insensitive)",
      idDis.status === "disagree" && idDis.by === "id" && idDis.code === 6, JSON.stringify(idDis));
    check("cross-check by route id: Illinois ids with inner double spaces match",
      cc([sid(3, 0, "016  20370 000000")], [sid(3, 5, "016 20370 000000")]).by === "id");
    check("cross-check: no shared id falls back to the distance rule",
      cc([sid(3, 0, "A")], [sid(7, 0, "B"), sid(3, 4, "C")]).by === "distance");
    check("cross-check: a missing id on this source falls back to the distance rule",
      cc([seg(3, 0)], [sid(3, 2, "A")]).by === "distance");
    const ap = (verdict, reason, c, name = "State DOT") => core.applyCrossCheck(verdict, reason, c, name);
    const dis = ap("Federal aid - Rural Other Principal Arterial", "", cc([seg(3, 25)], [seg(6, 25)]));
    check("disagree turns a red verdict into Review - Sources disagree",
      dis.verdict === "Review - Sources disagree" && dis.reason === "Sources disagree" && dis.note === "State DOT: Rural Minor Collector", JSON.stringify(dis));
    const dis2 = ap("Non-federal aid - Rural Minor Collector", "", cc([seg(6, 25)], [seg(3, 25)]), "HPMS");
    check("disagree turns a green verdict into Review - Sources disagree",
      dis2.verdict === "Review - Sources disagree" && dis2.note === "HPMS: Rural Other Principal Arterial", JSON.stringify(dis2));
    const kept = ap("Review - Nearby FHWA road", "Nearby FHWA road", cc([seg(7, 5), seg(3, 90)], [seg(3, 5)]));
    check("a row already under review keeps its reason and gains the note",
      kept.verdict === "Review - Nearby FHWA road" && kept.reason === "Nearby FHWA road" && kept.note === "State DOT: Rural Other Principal Arterial", JSON.stringify(kept));
    const tie = ap("Review - Conflicting classes", "Conflicting classes", cc([seg(3, 25), seg(6, 25)], [seg(6, 25)]));
    check("Conflicting classes row is told the other source's class",
      tie.verdict === "Review - Conflicting classes" && tie.note === "State DOT: Rural Minor Collector", JSON.stringify(tie));
    const same = ap("Federal aid - Urban Minor Arterial", "", cc([seg(4, 5)], [seg(3, 5)], true), "HPMS");
    check("class-differs leaves the verdict and notes the other class",
      same.verdict === "Federal aid - Urban Minor Arterial" && same.reason === "" && /^HPMS class: Other Principal Arterial/.test(same.note), JSON.stringify(same));
    const agree = ap("Federal aid - Urban Minor Arterial", "", cc([seg(4, 5)], [seg(4, 5)], true));
    check("agree changes nothing and adds no note", agree.verdict === "Federal aid - Urban Minor Arterial" && agree.note === "");
    const unclear = ap("Non-federal aid - Rural Local", "", cc([seg(7, 5)], [seg(0, 2)]));
    check("an unclear second source changes nothing", unclear.verdict === "Non-federal aid - Rural Local" && unclear.note === "");
  }
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
  check("road list dedups by name keeping nearest, prefers mixed-case, sorted",
    merged.length === 2 && merged[0].name === "Main St" && merged[0].distFt === 10 && merged[1].name === "Q Ave");
  // test-7.16 request 3: the same physical road written two ways (directional
  // prefix + street-type abbreviation + case) must collapse to ONE entry.
  const merged2 = core.mergeRoadList([
    { name: "MERIDIAN ST", distFt: 167 }, { name: "N Meridian St", distFt: 163 },
    { name: "HARRISON PKY", distFt: 230 }, { name: "Harrison Pkwy", distFt: 227 },
    { name: "W Market St", distFt: 197 }, { name: "E Market St", distFt: 226 }, { name: "MARKET ST", distFt: 199 }]);
  check("road list collapses directional/abbreviation/case duplicates to one road each",
    merged2.length === 3 &&
    merged2[0].name === "N Meridian St" && merged2[0].distFt === 163 &&
    merged2[1].name === "W Market St" && merged2[1].distFt === 197 &&
    merged2[2].name === "Harrison Pkwy" && merged2[2].distFt === 227,
    "got " + JSON.stringify(merged2.map(r => [r.name, r.distFt])));
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
