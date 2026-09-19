// Verifies the bottom-right pointer coordinate readout (the ArcGIS-style
// "wherever the pointer is" lat/lon box) and MEASURES what it costs, since
// the whole question about it was whether a per-mousemove readout would slow
// the map down.
//
// Correctness asserted:
//   - the control exists in Leaflet's bottomright corner, ABOVE the
//     attribution line, and starts on the map centre
//   - a mousemove at a known container point prints exactly what
//     map.containerPointToLatLng() says is under that point (5 dp)
//   - panning with the pointer held still re-projects the held point
//     (the geography under a stationary cursor changes during a pan)
//   - leaving the map falls back to the centre
//   - clicking the COORDINATE copies exactly what is displayed, and the
//     copied text parses back through parseCoordinates() in either format
//   - the small DD/DMS box beside it switches the format, labels itself
//     with the format in force, toggles from the keyboard too, and the
//     choice survives a reload
//
// Cost measured (printed, and budgeted loosely so this fails only on a real
// regression):
//   - per-event dispatch cost over the map vs the same burst dispatched off
//     the map (the delta is Leaflet's own container handler + this readout)
//   - DOM writes per burst: must be one per animation frame, NOT one per
//     event (that is the whole reason the handler defers to rAF)
//   - longest long task during a 4x-throttled pointer sweep at street zoom
//     with the real tiles rendering
//
//   cd build/web-tests && node verify-coord-readout.mjs
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
const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });
await context.grantPermissions(["clipboard-read", "clipboard-write"]);
const page = await context.newPage();
const pageErrors = [];
page.on("pageerror", e => pageErrors.push(e.message));
await page.goto(`http://127.0.0.1:${port}/index.html`, { waitUntil: "load" });
await page.waitForTimeout(1200);

const readout = () => page.evaluate(() => {
  const el = document.querySelector(".coordbox");
  return el && el.querySelector(".cr-val").textContent;
});

console.log("placement + initial state");
const place = await page.evaluate(() => {
  const el = document.querySelector(".coordwrap");
  if (!el) return null;
  const corner = el.parentElement;
  const attrib = document.querySelector(".leaflet-control-attribution");
  const r = el.getBoundingClientRect(), a = attrib && attrib.getBoundingClientRect();
  const legend = document.getElementById("siteLegend").getBoundingClientRect();
  return { corner: corner.className, sameCorner: !!attrib && attrib.parentElement === corner,
           aboveAttrib: !!a && r.bottom <= a.top + 1, clearsLegend: legend.bottom <= r.top + 1,
           inMap: r.right <= document.getElementById("map").getBoundingClientRect().right + 1 };
});
ok(!!place, "the readout control is on the map");
ok(place && /leaflet-bottom/.test(place.corner) && /leaflet-right/.test(place.corner), "sits in the bottom-RIGHT corner", place && place.corner);
ok(place && place.sameCorner && place.aboveAttrib, "stacks above the attribution line rather than over it");
ok(place && place.clearsLegend, "the site legend clears it (legend bottom is above the box)");
const role = await page.evaluate(() => {
  const box = document.querySelector(".coordbox"), btn = document.querySelector(".coordfmt");
  return { boxRole: box.getAttribute("role"), boxTab: box.tabIndex,
           btnRole: btn.getAttribute("role"), btnTab: btn.tabIndex, btnText: btn.textContent,
           labels: box.textContent.replace(/[-\d.,\u00b0'"\sNSEW]/g, ""),
           sideBySide: Math.abs(box.getBoundingClientRect().top - btn.getBoundingClientRect().top) < 2
                       && btn.getBoundingClientRect().left >= box.getBoundingClientRect().right - 1 };
});
ok(role.boxRole === "button" && role.boxTab === 0, "the coordinate is a real button (role + keyboard focusable)", `${role.boxRole}/${role.boxTab}`);
ok(role.btnRole === "button" && role.btnTab === 0, "so is the format switch", `${role.btnRole}/${role.btnTab}`);
ok(role.sideBySide, "the switch sits beside the coordinate, on the same line");
ok(role.labels === "", "the coordinate box carries no label text, just the number", JSON.stringify(role.labels));
ok(role.btnText === "DD", "the switch is labelled with the format in force", role.btnText);
const first = await readout();
ok(first && /^-?\d+\.\d{5}, -?\d+\.\d{5}$/.test(first), "opens on a decimal-degree coordinate", first);

console.log("pointer tracking");
const mapBox = await page.evaluate(() => {
  const r = document.getElementById("map").getBoundingClientRect();
  return { x: r.x, y: r.y, w: r.width, h: r.height };
});
const probe = async (cx, cy) => {
  await page.mouse.move(mapBox.x + cx, mapBox.y + cy);
  await page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
  const shown = await readout();
  const truth = await page.evaluate(([cx, cy]) => {
    const ll = map.containerPointToLatLng(L.point(cx, cy));
    return ll.lat.toFixed(5) + ", " + ll.lng.toFixed(5);
  }, [cx, cy]);
  return { shown, truth };
};
const p1 = await probe(400, 300);
ok(p1.shown === p1.truth, "prints the latlng Leaflet projects for that pixel", `${p1.shown} vs ${p1.truth}`);
const p2 = await probe(900, 620);
ok(p2.shown === p2.truth && p2.shown !== p1.shown, "follows the pointer to a second point", `${p2.shown} vs ${p2.truth}`);

console.log("pan / zoom with the pointer held still");
// Zoom in first: at the opening region view maxBoundsViscosity pins the map,
// so a pan there would move nothing and prove nothing.
await page.evaluate(() => map.setView([43.0389, -87.9065], 12, { animate: false }));
await probe(900, 620);
const before = await readout();
await page.evaluate(() => map.panBy([260, 180], { animate: false }));
await page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
const afterPan = await readout();
const panTruth = await page.evaluate(([cx, cy]) => {
  const ll = map.containerPointToLatLng(L.point(cx, cy));
  return ll.lat.toFixed(5) + ", " + ll.lng.toFixed(5);
}, [900, 620]);
ok(afterPan !== before, "the value changes when the map moves under a stationary pointer", `${before} -> ${afterPan}`);
ok(afterPan === panTruth, "and it is re-projected correctly, not stale", `${afterPan} vs ${panTruth}`);

console.log("leaving the map");
await page.mouse.move(mapBox.x - 40, mapBox.y + 20);   // over the left sidebar
await page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
const left = await readout();
const centre = await page.evaluate(() => {
  const c = map.getCenter();
  return c.lat.toFixed(5) + ", " + c.lng.toFixed(5);
});
ok(left === centre, "falls back to the map centre", `${left} vs ${centre}`);

console.log("click the coordinate to copy it");
await probe(500, 400);
const dd = await readout();
// Hovering means moving the pointer INTO the map container: if that move were
// tracked it would change the value before the click landed, so this also
// covers the freeze-over-controls rule.
await page.hover(".coordbox");
await page.waitForTimeout(120);
ok((await readout()) === dd, "the value holds while the pointer is over the box itself", await readout());
await page.click(".coordbox");
await page.waitForTimeout(150);
const clip = (await page.evaluate(() => navigator.clipboard.readText().catch(() => ""))).trim();
ok(clip === dd, "copies exactly what is displayed", `${JSON.stringify(clip)} vs ${dd}`);
ok(/copied/i.test(await page.evaluate(() => document.getElementById("toast").textContent)),
   "confirms the copy in the toast");
const ddPair = dd.split(",").map(Number);
const ddBack = await page.evaluate(t => { const p = parseCoordinates(t)[0]; return p && !p.invalid ? { lat: p.lat, lon: p.lon } : null; }, dd);
ok(ddBack && ddBack.lat === ddPair[0] && ddBack.lon === ddPair[1],
   "the copied text parses back through parseCoordinates", JSON.stringify(ddBack));

console.log("the DD/DMS switch");
await page.click(".coordfmt");
await page.waitForTimeout(120);
const dms = await readout();
ok(/^\d+°\d{2}'\d{2}\.\d"[NS] \d+°\d{2}'\d{2}\.\d"[EW]$/.test(dms), "one click switches the coordinate to DMS", dms);
ok((await page.evaluate(() => document.querySelector(".coordfmt").textContent)) === "DMS",
   "and the switch relabels itself");
const back = await page.evaluate(t => { const p = parseCoordinates(t)[0]; return p && !p.invalid ? { lat: p.lat, lon: p.lon } : null; }, dms);
ok(back && Math.abs(back.lat - ddPair[0]) < 0.0001 && Math.abs(back.lon - ddPair[1]) < 0.0001,
   "the DMS text parses back to the same point", `${JSON.stringify(back)} vs ${dd}`);
await page.click(".coordbox");
await page.waitForTimeout(150);
ok((await page.evaluate(() => navigator.clipboard.readText().catch(() => ""))).trim() === dms,
   "copying still copies what is on screen, now in DMS");
await page.click(".coordfmt");
await page.waitForTimeout(120);
ok((await readout()) === dd, "a second click switches back to decimal degrees", await readout());
await page.focus(".coordfmt");
await page.keyboard.press("Enter");
await page.waitForTimeout(120);
ok(/[NS]/.test(await readout()), "Enter on the focused switch toggles it too", await readout());

console.log("the choice sticks");
await page.reload({ waitUntil: "load" });
await page.waitForTimeout(1200);
ok(/[NS]/.test(await readout()), "DMS survives a reload", await readout());
ok((await page.evaluate(() => document.querySelector(".coordfmt").textContent)) === "DMS",
   "and the switch comes back labelled DMS");
await page.click(".coordfmt");                      // back to DD for the cost legs
await page.waitForTimeout(120);
ok(/^-?\d+\.\d{5}, -?\d+\.\d{5}$/.test(await readout()), "and switches back off again", await readout());
const fits = await page.evaluate(async () => {
  // Measure BOTH formats: DMS is the longer string, and the box is fixed-size.
  const el = document.querySelector(".coordbox"), btn = document.querySelector(".coordfmt");
  const read = () => ({ sw: el.scrollWidth, cw: el.clientWidth, sh: el.scrollHeight, ch: el.clientHeight,
                        bw: btn.scrollWidth, bcw: btn.clientWidth, t: el.textContent });
  const a = read();
  btn.click();
  await new Promise(r => requestAnimationFrame(r));
  const b = read();
  btn.click();
  await new Promise(r => requestAnimationFrame(r));
  return [a, b];
});
ok(fits.every(f => f.sw <= f.cw && f.sh <= f.ch && f.bw <= f.bcw), "neither fixed-size box clips its text",
   fits.map(f => `${f.t.trim()} ${f.sw}/${f.cw}`).join(" | "));

console.log("cost: per-event dispatch");
// Dispatch identical synthetic mousemove bursts at the map container and at
// a bare off-map element. The delta is everything the map does per pointer
// event (Leaflet's own container handler + this readout).
const BURST = 4000;
const burst = await page.evaluate(async n => {
  const mapEl = document.getElementById("map");
  const off = document.createElement("div");
  document.body.appendChild(off);
  const r = mapEl.getBoundingClientRect();
  const evs = [];
  for (let i = 0; i < n; i++) {
    const x = r.x + 60 + (i * 7) % (r.width - 120), y = r.y + 60 + (i * 11) % (r.height - 120);
    evs.push({ x, y });
  }
  const run = el => {
    const t0 = performance.now();
    for (const e of evs) el.dispatchEvent(new MouseEvent("mousemove", { clientX: e.x, clientY: e.y, bubbles: true }));
    return performance.now() - t0;
  };
  run(mapEl); run(off);                       // warm both paths
  const onMap = run(mapEl), offMap = run(off);
  off.remove();
  return { onMap, offMap };
}, BURST);
const perEventUs = (burst.onMap - burst.offMap) / BURST * 1000;
console.log(`      ${BURST} synthetic mousemoves: ${burst.onMap.toFixed(1)} ms on the map, ` +
            `${burst.offMap.toFixed(1)} ms off it -> ${perEventUs.toFixed(2)} us/event of map+readout work`);
ok(perEventUs < 25, "per-event cost stays microseconds, not milliseconds", `${perEventUs.toFixed(2)} us`);

// The readout's OWN share of that: the two projections it runs per event
// (the rest of the delta is Leaflet's own container handler, which predates
// this feature and runs whether or not anything listens).
const ownUs = await page.evaluate(n => {
  const mapEl = document.getElementById("map"), r = mapEl.getBoundingClientRect();
  const evs = [];
  for (let i = 0; i < n; i++)
    evs.push(new MouseEvent("mousemove", { clientX: r.x + 60 + (i * 7) % (r.width - 120), clientY: r.y + 60 + (i * 11) % (r.height - 120) }));
  const run = () => {
    const t0 = performance.now();
    for (const e of evs) { const p = map.mouseEventToContainerPoint(e); map.containerPointToLatLng(p); }
    return performance.now() - t0;
  };
  run();
  return run() / n * 1000;
}, BURST);
console.log(`      the readout's own work (2 projections/event): ${ownUs.toFixed(2)} us/event`);
ok(ownUs < 10, "the readout's own per-event work is under 10 us", `${ownUs.toFixed(2)} us`);

console.log("cost: DOM writes are per FRAME, not per event");
const writes = await page.evaluate(async n => {
  const val = document.querySelector(".coordbox .cr-val");
  let n_writes = 0;
  const obs = new MutationObserver(m => { n_writes += m.length; });
  obs.observe(val, { characterData: true, childList: true, subtree: true });
  const mapEl = document.getElementById("map"), r = mapEl.getBoundingClientRect();
  for (let i = 0; i < n; i++)
    mapEl.dispatchEvent(new MouseEvent("mousemove", { clientX: r.x + 60 + i % 300, clientY: r.y + 60 + i % 200, bubbles: true }));
  await new Promise(res => requestAnimationFrame(() => requestAnimationFrame(res)));
  obs.disconnect();
  return n_writes;
}, 600);
console.log(`      600 events in one frame -> ${writes} DOM write(s)`);
ok(writes <= 2, "600 events in a frame coalesce into at most one paint", String(writes));

console.log("cost: throttled pointer sweep at street zoom, real tiles");
const client = await page.context().newCDPSession(page);
await client.send("Emulation.setCPUThrottlingRate", { rate: 4 });
await page.evaluate(() => map.setView([43.0389, -87.9065], 16, { animate: false }));
await page.waitForTimeout(2500);
await page.evaluate(() => {
  window.__lt = [];
  new PerformanceObserver(l => { for (const e of l.getEntries()) window.__lt.push(Math.round(e.duration)); })
    .observe({ type: "longtask", buffered: false });
});
for (let i = 0; i < 60; i++) {
  await page.mouse.move(mapBox.x + 100 + (i * 17) % 900, mapBox.y + 80 + (i * 23) % 600);
  await page.waitForTimeout(16);
}
await page.waitForTimeout(400);
const lt = await page.evaluate(() => window.__lt.slice().sort((a, b) => b - a));
console.log(`      longest long tasks during the sweep (4x throttle): ${lt.slice(0, 5).join(", ") || "none"} ms`);
// Budget is loose on purpose: anything long that shows up here is the tile
// renderer catching up under 4x throttling, not the readout — its own share
// is the microseconds measured above. This leg is a regression tripwire for
// "the readout made pointer movement janky", not a tile-rendering budget.
ok(!lt.length || lt[0] < 500, "no long task over 500 ms while tracking the pointer", lt.length ? lt[0] + " ms" : "none");
await client.send("Emulation.setCPUThrottlingRate", { rate: 1 });

ok(pageErrors.length === 0, "no page errors", pageErrors.join(" | "));

await browser.close();
server.close();
console.log(failures ? `\n${failures} check(s) FAILED` : "\nall checks passed");
process.exit(failures ? 1 : 0);
