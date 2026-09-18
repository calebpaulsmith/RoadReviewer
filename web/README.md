# FHWA Road Checker (provisional title) — web prototype

A single static page that performs RoadReviewer's Workflow 1 (Classify
Roads) in the browser. Paste coordinates into the textarea and each point
is parsed, classified, and pinned on the map automatically — no submit
button, no server, nothing stored. The page is a direct JavaScript port of
`src/modClassify.bas` + `src/modConstants.bas` + `src/modHttp.bas`'s
distance math: same layer URLs, same retired-segment filters, same
exact-intersect-then-buffer fallback (the **Detection Buffer** control,
default 250 ft, matching the Excel search buffer), same per-road
distances, and the same PR #24 verdict model — the **closest** road
segment decides red vs blue, amber only downgrades blue with an
explicit Review Reason ("Second road close" / "Nearby FHWA road" /
"Urban boundary edge"), and red never downgrades. Same red/blue/amber
buckets as the Sites table and KML export. A "Download PDF Report"
button turns the classified points into a citeable PDF — see below.

## Results pane

Results render as one compact card per site (not a spreadsheet-style
table), so nothing ever scrolls horizontally and many sites fit on
screen. Each card carries a solid-color verdict badge (FEDERAL AID /
NON-FEDERAL AID / REVIEW / FAILED) over a matching row tint and left
border, then unpacks the Excel tool's pipe-joined class string into
individual road chips: "N road segments within {radius} ft, nearest
first:" followed by one chip per distinct road/class **with its measured
distance** ("Local (0 ft)", "Major Collector (19 ft)"), each swatched in
the standard FHWA class color (the same palette the map overlay and PDF
figures use, so chip color = line color on the map). The first chip is
tagged **closest** — that's the segment that decided red vs blue. Where
the source layer carries a road name on the class feature (both
Wisconsin layers), the chip pairs name + class ("STH 86 E · Major
Collector"); Michigan/Indiana class layers publish no name field, so
their chips show class only (with a ×n multiplier for repeated
segments). A "Roads:" line merges every named road (state route layers +
Census TIGER) nearest-first with distances — the Excel Road Name column.
An Urban/Rural chip cites the ACUB polygon name (including the
"Rural · edge of …" boundary-ambiguous case). Yellow REVIEW verdicts show
their ≤3-word reason inline, with a plain-language explanation in the
tooltip and marker popup. Per-row links: **ArcGIS map** (Map Viewer
pinned on the point with the state class layer — parity with the Excel
"AGOL NFC Layer" column), **Public map** (the state's official app at
its root — pair it with **Download GeoJSON** and drag the file onto the
map to see your sites there), Google, Street View, FIRMette, Source.
CSV / copy-for-Excel exports keep the flat columns (now including Review
Reason and the distance-annotated road list); Download GeoJSON emits the
same FeatureCollection shape as the Excel tool's GeoJSON export
(verdict, color, class, roads as properties).

## Site-by-site review

Pasted site names are surfaced as labels on the map pins (permanent up to
25 sites, hover past that). Clicking a result card or a pin zooms straight
to that site; **Prev / Next** above the map step through the sites one by
one (wrapping at the ends). While a site is selected and "Site layers"
is checked, the page fetches the authoritative geometry around it — the
state's road-class layer(s) and the 2020 ACUB urban-boundary polygon,
via the same frame-envelope queries the PDF figures use — and draws it
directly on the interactive map in each source's own published renderer
colors. An on-map legend names the exact layers drawn, lists only the
classes present, links each layer to a live ArcGIS view centered on the
site, and links to the citations page below. Per-site layer fetches are
cached, so stepping back and forth doesn't re-query.

## Full-page map + live layer mirror (2026-09-14)

The map is the page hero: it fills the viewport width and most of its
height, with the coordinate-input controls in a collapsible card floating
over its top-left corner ("☰ Input panel" toggles it) and the site legend
floating bottom-right. On phone-width or short viewports the panel drops
back into normal flow above the map.

The **"Live class + urban layers"** toggle (on by default) continuously
mirrors every wired state's official functional-class layer and the NTAD
2020 ACUB urban boundaries across whatever area is visible — pan or zoom
anywhere in the six Region V states and the page fetches that viewport's
geometry from the same public services the classifier queries, drawn in
each source's own published renderer colors (states are picked by
bounding-box intersection with the view, so a border viewport draws both
sides). A bottom-left legend lists exactly the classes present in view.

Two constraints shape it, both disclosed in the legend:

- **Zoom gating — progressive by class (2026-09-14).** The services cap a
  single query at ~1000–2000 features and the AGOL-hosted FeatureServers
  are *Query-only* (no server-side `/export` rendering), so the mirror
  displays scale-dependently the way professional viewers do: **principal
  arterials appear from zoom 10, minor arterials from zoom 12, and the
  full network (collectors + local streets) from zoom 13** — at the lower
  bands each state is queried with a server-side class filter
  (arterials are a tiny fraction of the segments: metro Columbus at a
  z11-sized view is 23,768 segments total but 1,144 at class ≤ 4,
  live-verified), with geometry generalized to the pixel grid. The
  urban-boundary polygons load from zoom 7 the same way. The legend
  always says what's shown ("principal arterials and up — zoom in for
  collectors and local streets"); if a query still hits the record cap
  (downtown Chicago's fine-chopped arterials can), the legend discloses
  the truncation. Per-state filter syntax differs (numeric vs string
  codes; WisDOT's local layer encodes urban/rural into its code) — all
  six confirmed live via `returnCountOnly`.

  This progressive live path is now the **fallback** — see the next
  section: where a state's pre-built HPMS tileset exists, the class
  display comes from static tiles instead and none of these live layer
  queries fire.

## Full-page app shell + offline basemap, locked to Region V (2026-09-14)

Per user direction, the map now IS the page: no page scrolling, and a
full-height left pane carries everything else — title, inputs (the
coordinates box's grey placeholder is real example lines and clears on
click), Find box, the results list, exports, and the transparency
panels. The old header and "Private by design" banner are gone (the
privacy note lives at the pane's foot). Prev/Next/Plot-all and the layer
toggles float in a bar over the map's top-left.

- **Locked to Region V.** `maxBounds` covers the six states and
  `minZoom` is pinned to the region-fit zoom (recomputed on resize,
  floor z6) — you cannot zoom or pan away from MI/IN/WI/MN/IL/OH.
- **The road basemap is served from this site** —
  `web/tiles/basemap.pmtiles`, a 50 MB Protomaps/OSM extract of the six
  states (z0-11, layers earth/water/roads/boundaries/places), rendered
  by the vendored protomaps-leaflet **5.1.0** `light` *flavor*. Browsing
  the road map fetches nothing from a tile CDN. **Renderer/tile schema
  must match (bug fixed 2026-09-14):** the daily Protomaps builds are the
  v4 schema (property `kind`; rivers are LINE features inside the water
  layer). The 3.x renderer first vendored only knew v3's `pmap:kind`, so
  roads/boundaries/place labels never matched a rule (blank grey land at
  every zoom) and every river got filled as a polygon (long cyan "sliver"
  triangles across Minnesota). Any future protomaps-leaflet upgrade must
  stay on a version whose flavors read the schema version of
  `basemap.pmtiles`.
- **Atlas palette (2026-09-15, per user: "a beautiful map, something
  you'd expect when opening a map product, not a bunch of purple
  lines").** The basemap renders through a custom flavor object
  (`ATLAS_FLAVOR` in index.html, handed to protomaps-leaflet's exported
  `paintRules`/`labelRules` because the `flavor` option only takes the
  built-in names): warm ivory paper `#f3ecd9`, muted blue-green water
  `#a8c6c7`, sage parks/woods, ochre highways, restrained minor roads,
  brown-grey type; the street-level OpenFreeMap detail and the region
  mask use the same `ATLAS` swatches. The OPENING map is the atlas
  alone: class lines and urban boundaries now start at **z9**
  (`BROWSE_CLASS_ZOOM` / `BROWSE_ACUB_ZOOM`), even though the class
  tiles carry interstates from z6.
- **Cut to the state shapes.** `web/data/r5-states.geojson` (Census
  TIGERweb state boundaries, generalized, 58 KB) is drawn as one mask
  polygon — world outer ring, the six states as holes — in a pane above
  the basemap + detail tiles and below the class/ACUB overlays, filled
  with the page ground colour; the hole edges are the state borders. The
  basemap extract itself is still six rectangles (no local `pmtiles`
  CLI to re-extract by polygon), but nothing outside the states shows. **Satellite stays live Esri**,
  fetched only when switched to; the live Esri street layer remains as
  the automatic fallback when the basemap file is missing (or on
  `file://`). Build pipeline + gotchas: `build/tiles/README.md`.
- **Compact result rows.** Each row shows one line — site name, coords,
  state, verdict badge — and clicking it both selects the site on the
  map and expands the detail (chips, roads, urban area, links).
  **⧉ View and Export** (results header) and the blue **View and Export →**
  button pinned to the bottom of the pane both EXPAND the pane over the map
  (2026-09-17; see "View and Export" below) — there is no pop-out any more.
- **Live street-level detail under the roads** (per user direction:
  hosted road map, "the other stuff popped in live underneath"):
  buildings, parks/land use, street names, house numbers and — the
  Google-style ground — **white road ribbons with gray casings**, so
  the FHWA class lines read as colored centerlines ON the streets. All
  from **OpenFreeMap** (keyless, no-limit public OSM vector tiles,
  CORS `*` — TileJSON at `https://tiles.openfreemap.org/planet`
  resolves the versioned `{z}/{x}/{y}.pbf` template at runtime). It
  draws in its own pane between the offline basemap (z200) and the
  HPMS class overlay (z350), fetch-gated to z13+ so region browsing
  stays fully offline; labels appear z14+ (street names), z18+ (house
  numbers — z17 was wall-to-wall numbers downtown). OpenFreeMap's
  `water`/`waterway` are redrawn here too (same colour as the basemap's
  light flavor): the offline basemap stops at z11 and its lakes
  overzoomed 16x were blocky at street zooms. Buildings from z15. **No POI markers** — the first cut drew POI dots + names
  and the user rejected the clutter. Best-effort: blocked or down, the
  map just shows no detail layer.

## Two input tabs: Coordinate Input + Search & Collect (2026-09-14)

Per user direction the left pane's input area is two tabs:

- **Coordinate Input** — the paste-a-batch flow: the Search **buffer**
  select (renamed from "Search radius") and the coordinates box. The
  State dropdown is GONE — the state is always auto-detected per point.
- **Search** — since 2026-09-17 there are no tabs: the search box sits at
  the top of the one pane, above the coordinates box, with no Find button
  (results appear as you type, grouped under State / County / City /
  Township / Road headings; cities come from TIGERweb's incorporated-places
  layer, which the earlier search never queried). Clicking a county, city or
  township zooms there AND makes it the **default area** (blue chip, ✕ to
  clear, remembered in this browser): addresses typed without a city/state
  are geocoded inside it, bare road names (`Q Ave`) are looked up inside it,
  and the road-name search runs inside it. Anything ambiguous — several
  geocoder matches, several separate stretches of a road with that name —
  becomes a picker on the row. Right-clicking a pin offers **Move pin**
  (drag; the line and verdict follow) and **Delete pin**. The add-a-point form is gone (2026-09-17, per user: it
  repeated the coordinates box). Sites are added with the **pin button** at
  the top right of the map: one click arms it for ONE pin (button turns
  blue, the pointer becomes a pin), a second click keeps it on (orange, ∞)
  for dropping many, a third — or Escape — turns it off; scroll-zoom and
  drag still navigate. A map click drops a bouncing pin and appends
  `Point N, lat, lon` to the coordinates box, so the site goes through the
  normal parse → classify flow; the pin's label opens as a name box with
  the default name selected (the same "Point N" rule as an unnamed pasted
  line) — just type to replace it, Enter or clicking elsewhere finishes,
  typing is optional. The map does not jump while you name the pin. Points
  saved by the old form still load from localStorage (their rows keep the
  "remove" link) but nothing writes there any more.
- **Exports**: CSV/Copy-for-Excel/GeoJSON gained a **Note** column, and
  a new **KMZ** export (zipped KML via the FIRMette bundle's store-zip
  builder) writes red/blue/yellow pushpins by verdict with name,
  status, class, urban area, roads and note in each placemark — the
  same conventions as the Excel tool's KML.

## Accepted coordinate formats (2026-09-15)

The paste box runs `parseCoordinates` (a dropped pin writes a
`name, lat, lon` line into that same box), so it accepts the following:

- **Decimal degrees**, in either order, with or without a site name, in any
  mix of commas, tabs and spaces: `42.28536, -85.57025` ·
  `Culvert on Q Ave⇥42.6911⇥-84.5360` · `39.9876⇥-86.0128⇥CR 550 N`.
- **Degrees / minutes / seconds and degrees / decimal-minutes**, which is what
  handheld GPS units and FEMA paperwork usually carry:
  `42°17'07.3"N 85°34'12.9"W` · `N42°17'07.3" W85°34'12.9"` ·
  `42 17 07.3 N, 85 34 12.9 W` · `42 17 07.3, -85 34 12.9` ·
  `42° 17.122' N, 85° 34.215' W`. Hemisphere letters, a leading minus, or
  neither (an unsigned longitude over a Region V latitude is read as west).

DMS is parsed **before** the decimal scan, and that ordering is the point: the
decimal scan takes the last in-range number PAIR on the line, so
`42 17 07.3, -85 34 12.9` used to come back as 34, -85 — a confident
federal-aid verdict for a point in Alabama, with nothing on screen to say it
was wrong. Symbol-bearing DMS was merely rejected; the signed forms were the
dangerous ones.

Degrees must be whole and minutes/seconds at most two digits, and a
hemisphere letter only counts when it stands alone — otherwise a street
number, a ZIP, or the "e" in `Culvert on Q Ave` would pose as part of a
coordinate. A name *ending* in a direction word (`Rose Drive W 42°17'07.3"N
…`) still hands that letter to the coordinate and flips the latitude
negative, so when the first read isn't a Region V coordinate the leading
hemisphere is dropped and the line re-matched. `build/verify-web-core.mjs` pins every form above (each one is
the §4.2 Kalamazoo test point written differently) plus the name cases.

## Addresses and road names in the same box (2026-09-15)

A line with no coordinate is sorted **locally** — nothing is sent to sort it —
into one of three kinds, and the box highlights the ones that need a decision:

| you type | kind | what happens |
|---|---|---|
| `42.28536, -85.57025` | coordinate | classifies as always |
| `Portage Rd, Portage MI` · `CR 550 N, Hamilton County IN` | road | looked up **automatically** in Census TIGERweb and the **whole road** classified |
| `5201 Portage Rd, Portage MI 49002` | address | highlighted; **waits for the one `Geocode N addresses` button** |
| `Q Ave` (a road with no place) · anything else | unknown | shown as unreadable — a road with no place would mean a six-state search |

**Road lines.** The place resolves to a county, township, incorporated place
or CDP (by `BASENAME`, scoped to the typed state or to the six states — more
than one hit with no state is reported as ambiguous), then the road's edges
are fetched by street *components* inside that extent. Up to 12 edges
(longest first) are classified at their midpoints: one class and one verdict
bucket → that verdict; otherwise `Review – Mixed classes on road`, with a
class-by-length chip (`Local 2.1 mi · Major Collector 0.4 mi`). The pin is
the longest edge's midpoint and the edges draw on the map. TIGERweb sends
CORS headers, so this is an ordinary `fetch()` — no geocoder involved.

**Address lines.** The Census one-line geocoder sends **no
`Access-Control-Allow-Origin` header** (verified 2026-09-15 against two
controls: `services.arcgis.com` sends `*`, `tigerweb.geo.census.gov` reflects
the origin, the geocoder sends only `Vary: Origin`), so the page can't
`fetch()` it. It does answer `format=jsonp`, which means a remote script
executing — so each request runs inside a **throwaway sandboxed iframe**:
`sandbox="allow-scripts"` only (opaque origin: no access to this page's DOM,
localStorage or cookies), the https host fixed in the frame's own source, a
per-request id that is also the callback name, `postMessage` accepted only
from that frame's window with origin `"null"` and that id, the reply reduced
to whitelisted, range-checked fields before any of it is used, and the frame
removed on success, error or a 20 s timeout. **Nothing is sent on keystroke,
paste, blur or parse.** The button is the network action; it geocodes every
unresolved address (three at a time) and disappears when none are left.

**Why the street name is the anchor and the geocoded point only a hint.** The
geocoder returns a TIGER address-range *interpolation* — a point on the
centerline, nudged to one side. Measured 2026-09-15, `5201 Portage Rd` landed
**20 ft from Airview Blvd and 21 ft from Portage Rd**, so the closest-road
model would have picked the wrong street by one foot. Instead the geocoder's
parsed street components (`streetName` / `suffixType` / directions) are
matched against the Census roads within 120 m — base name must agree, type
and directions add or subtract — the point is **snapped onto the matching
edge**, and the classification runs at the snap **and 150 ft each way along
that edge**: all agree → verdict; a class change inside that window →
`Review – Class change nearby`. No matching street → the raw point classifies
with `Review – Street not matched`. No geocoder match, several matches more
than 200 ft apart, or an unreachable geocoder → the line stays an unresolved
row with the reason, and the button offers it again.

**Typed coordinates are authoritative.** Geocoding never rewrites the box.
Exports carry the snapped coordinates under the geocoder's matched address;
the row shows a `from address · snapped to Portage Rd` chip (or `street not
matched`). Results are held in memory only, keyed by the line's text: edit a
line and it is unresolved again; reload and it takes another click.

## Cached-tile classification (2026-09-14) — verdicts from the hosted data

Per user direction ("couldn't we just check every single point with the
data we have cached?"), classification now runs against the tilesets
hosted with the site by default: road segments + FHWA classes from the
per-state HPMS tiles, urban/rural + urban-area name from the new
`tiles/acub.pmtiles` (549 Region V polygons of the 2020 Adjusted Urban
Areas, 3 MB). The SAME `computeVerdict` logic as the live path — only
the data source changes — so red/blue/amber rules are identical.
Instant, and zero per-point queries to the six state DOT servers or
NTAD; only the non-fatal Census TIGER street-name backfill stays live.
Rows show a "Source: FHWA HPMS 2024 tiles" chip (live rows show
"Source: MDOT live layer" etc.; every chip links to the source's public
product page, and the map legend's rows do the same on hover/click); a
**Live verdicts** checkbox under
"Data service URLs" restores the per-point live queries (slower, but
reflects reclassifications newer than the tileset year). The live path
also remains the automatic fallback (file://, missing tileset, read
error, out-of-region point).

The Search & Collect find box lost its state dropdown — searches cover
all six Region V states, and typing a state's name matches it directly.
Showing a searched road now also reports how many of your points lie
within the search buffer of it.

## Baked HPMS class tiles (2026-09-14) — the primary class display

Per user direction ("I don't love all this live querying — download the
HPMS data; high classes never change"), the road-class map layer is
served from **pre-built vector tiles** of FHWA's HPMS full-extent data —
every public road, class 1-7, locals included — one PMTiles file per
state in `web/tiles/`, built by `build/tiles/` (see its README for the
pipeline: quadtree envelope harvest of the BTS `HPMS_National_Current`
service → tippecanoe). Division of labor:

- **Display = tiles.** Instant at any zoom, full network including the
  small roads most reviewed points sit on, zero live queries, no record
  caps, works even when a state server is down. Progressive display is
  baked into the tileset (interstates z6 → locals z12). The legend names
  the HPMS year and which states are tile-served.
- **Verdicts = live.** Classification of pasted points still queries the
  state DOT's authoritative layer per point, exactly as before — that's
  where currency matters (collector/local reclassifications move the
  federal-aid line) and it's only a few queries per site.
- **Fallback = the live mirror above.** A state with no tileset (or a
  failed tile fetch, or a `file://` open) automatically keeps the live
  progressive class display.

**GitHub Pages still hosts everything.** PMTiles are read via HTTP range
requests, which Pages' CDN serves; the constraint is GitHub's 100 MB
per-file limit — per-state sizes are recorded in `build/tiles/README.md`
and a state that outgrows it splits into two files. Renderer:
`protomaps-leaflet` + `pmtiles` vendored in `web/vendor/` (no CDN),
drawing into the same canvas pane under the site pins.
- **Rendering.** Everything draws into one Leaflet `<canvas>` pane
  beneath the site pins and the per-site overlay, so a few thousand
  segments render without the per-element cost of SVG. Viewport fetches
  are debounced (250 ms), keyed-cached, and sequence-guarded so a stale
  response never paints over a newer view; failed fetches are not cached
  and retry on the next pan.

## Find on map + row filter (2026-09-14)

The **"Find on map"** box in the input panel searches by state, county,
township, or road name, via Census TIGERweb (free, no auth — already this
tool's street-name source; layers documented on the sources page).
**Suggestions appear as you type** (debounced at 2+ characters; fetches
are URL-cached so backspacing through a term costs nothing; Enter or the
Find button still search immediately):

- **State** — pick one from the dropdown and Find with the box empty to
  zoom straight to it (the dropdown also scopes every other search).
- **County / township** — type part of the name; counties (TIGERweb
  `State_County` layer 1) and county subdivisions (`Places_CouSub…`
  layer 1) are searched statewide, matches listed with their kind.
  Clicking one zooms to the boundary and draws it as a dashed outline
  (fetched generalized to the zoom — navigation cartography, not
  authoritative jurisdiction lines).
- **Road name** — searched across TIGER's full-detail Primary /
  Secondary / Local road layers, but only **within the visible map
  area** (a statewide un-indexed `LIKE` would be slow for everyone), so
  the flow is find-the-county-then-the-road; below zoom 11 the results
  say so. Matches group segments by full name; clicking one highlights
  the segments and zooms to them. Each road suggestion is annotated
  (asynchronously, so the list never waits) with the state's **FHWA
  functional class**, swatched in the standard class color — looked up
  from the state's own class layer at the midpoint of the road's longest
  matched segment via the same per-state query the classifier uses,
  closest segment wins, cached per road.

Site pins keep taking the row's verdict color (red / blue / amber, the
same buckets as the row tint and KML pushpins) once a pasted point
classifies; they got a bolder white ring and slightly larger radius so
the verdict reads clearly over the class-colored live road lines.

The results list gets a matching **row filter**: a text box that hides
result cards whose text doesn't match (site name, road/street names,
verdict, class, urban area) **and hides those sites' map pins with
them**, so the table and the map never show different sets. Beside it,
a **"filter by map view"** checkbox (off by default) additionally keeps
only the sites inside the current map bounds — Find a county or
township first and the list follows it, live as you pan. Zooming to a
single site (a row click, Prev/Next, or adding a collected point) is
deliberately exempt from that: it would otherwise collapse the list to
the one site you just clicked, so those moves re-apply the filter
against the last view the user chose. All of it is purely visual —
exports, the View and Export table, Prev/Next stepping and classification
always cover every site.

## Verdict colours (2026-09-17)

The verdicts are **red** (federal aid, `#cc3311`), **amber** (needs review,
`#ee9900`) and **blue** (non-federal aid, `#0077bb`) — not green. The old
green/yellow pair failed a colour-vision check (ΔE 4.4 for protan viewers);
this trio passes every check with a worst pair of ΔE 19, and every place the
colour appears (row tint, badge, pin, chart bar, PDF figure, GeoJSON
`VerdictColor`) also carries the words. The KMZ uses Google Earth's blue
pushpin for non-federal aid. The Excel workbooks still tint green/yellow.

## Quick Export + View and Export (2026-09-17)

**Quick Export ▾** (with the coordinate input, not under Auto-Detect) is a
plain dropdown: copy site + coordinates, copy Auto-Detect results, and one
DIRECT download per format (CSV, KMZ, GeoJSON, PDF report, FIRMettes ZIP)
— no dialogue in between. **View and Export →**, pinned to the bottom of
the pane (and "⧉ View and Export" in the results header), stretches the
left pane across the screen with a short slide; the map keeps a strip on
the right (about a sixth of the width) and zooms to fit every site, or
stays on the selected site if one is selected. The strip keeps Prev/Next.
Inside is the same format-tabbed, editable table that used to be the
modal dialogue (everything below still applies): clicking a row selects
that site exactly like the small list does — blue outline on the row, the
list row highlighted, the map zoomed right in — and **← Back to map** (or
Escape) restores the layout, applying any pending edit.

## Export tabs (2026-09-15)

The format tabs and the editable table were built as a modal dialogue and
now live inside View and Export; the mechanics are unchanged:

- **Tabs per format** — Excel, Google Earth (KMZ), GeoJSON, PDF — each
  with its own actions, its own options, and a **live preview** of what
  that format will actually write.
- **One editable table** under the tabs, showing every column the exports
  carry, **Note included**. Every cell is editable; edits are keyed by the
  site's coordinates (so they survive the re-render each keystroke in the
  coordinates box triggers) and flow into the CSV, the clipboard copies,
  the KMZ and the GeoJSON. A Note edit also sticks to a collected point in
  this browser. The PDF report re-queries the live layers to draw its
  figures, so it is deliberately not edit-driven.
- **Site Name, Latitude and Longitude write back.** Those three are the
  site's identity rather than export decoration, so editing one rewrites
  that site's line in the coordinates box
  — splicing just the number you changed, leaving your own separators and
  the other number as typed — or its collected record, and the site
  re-classifies at the new location with its pin moving to match. So a
  corrected coordinate gets you the verdict for the corrected spot, and the
  CSV, KMZ and GeoJSON can't disagree about where a site is. A latitude or
  longitude outside the range the coordinate parser itself accepts is
  refused and the cell reverts. The edit is saved as soon as you leave the
  cell, but the write-back and the re-check wait until you leave the **row**
  — tabbing from Latitude to Longitude is one correction, so a transposed
  pair is re-checked once rather than twice (the first time at a
  half-corrected spot). Enter, clicking outside the row, or closing the
  dialogue applies immediately; while the re-check runs, the dialogue's
  export buttons are held so a download can't omit the site.
- **Two clipboard actions** — *Copy site + coordinates* (just the name and
  lat/lon, in the layout the coordinates box itself accepts) and *Copy
  Auto-Detect results* (every column). Both confirm with a short toast;
  the page uses `navigator.clipboard` where the browser allows it (the
  live https site) and falls back to a hidden textarea, telling you to
  copy the preview by hand if a locked-down browser blocks both.
- **PDF map width** moved into the PDF tab. Its labels are plain feet then
  miles; the underlying values are unchanged metre half-widths, so the
  figures render exactly as before. *Open item: what this control should
  be is up for review — see CLAUDE.md §7b.*

## Data sources page

`sources.html` (linked from the header, the review legend, and every
result row's "Source" link) documents every layer the tool queries, by
state: organization, service URL, exact layer names, the fields read,
and the schema quirks that shaped the implementation (retired-segment
filters, Indiana's `record_status` domain and single-symbol renderer,
Wisconsin's embedded urban/rural category codes, the ACUB buffer floor,
and so on). Row/legend links anchor to the right state's section.

## FIRMette batch download (ZIP)

"Download FIRMettes (ZIP)" generates a FEMA FIRMette (flood-map extract
PDF) for every pasted site by driving FEMA's own Print FIRMette
geoprocessing service from the browser (submitJob → poll → download —
the same flow the Excel tool and FEMA's Map Service Center portal use;
CORS on every step confirmed live 2026-07-03), then bundles the PDFs
into one ZIP assembled in the browser by a small dependency-free
STORE-only writer (PDFs are already internally compressed). Batches are
capped at 20 sites per run because FEMA renders each PDF fresh (~1 MB,
seconds to a couple of minutes each); failures are reported per site and
don't sink the rest of the batch.

## PDF report

Click **Download PDF Report** to generate a PDF with a cover page (a
summary table of every classified site) followed by one page per site.
Each site page has **one combined, page-filling map**: the 2020 Adjusted
Census Urban Boundary (ACUB) polygon drawn underneath, and the state's
road functional-class polylines (MDOT/INDOT/WisDOT; for Wisconsin, both
the state-trunk and local-roads layers) drawn on top — so a single figure
answers both halves of the federal-aid question at once.

The map's coverage is set by the **PDF map width** dropdown next to the
button (Close ~0.4 mi / Standard ~0.75 mi / Wide ~1.5 mi / Very wide
~3 mi across; Standard is the default). The figure is drawn from fresh
geometry queries covering that whole frame — not just the segment that
produced the verdict — so the site appears in the context of the
surrounding road network. Below the figure, each page carries **clickable
source links** in two tiers. The **primary reference** (first, bold) is
the state's own official public map — Michigan's MDOT "NFC, NHS & ACUB"
ArcGIS Experience app, Indiana's INDOT "Functional Classification & Urban
Area Boundary" viewer — both showing functional class *and* the urban
boundary by default. It opens the app at its default extent (its
coordinate deep-link proved unreliable and was dropped). The **pinned
links** that follow open the FEMA-hosted ArcGIS Map Viewer centered and
markered on the exact site (MI's curated NFC/ACUB webmap; the FeatureServer
side-loaded for other states / the ACUB layer) — these are what put you on
your point. Wisconsin has no statewide interactive app (static PDFs only),
so it gets only the pinned link; MN/IL/OH link their DOT map portals
(EMMA / Getting Around Illinois / TIMS) at their roots. Plus a Google Maps
link. (Layer on/off
can't be driven through a URL in Experience Builder, but the wanted layers
are the app defaults — see `sources.html` for the per-state specifics.)

Under the data layers sits an **Esri World Street Map basemap** (roads +
street names for orientation), composited from tiles fetched for the
frame at report time. The tiles are loaded with
`crossOrigin="anonymous"` against Esri's CORS-enabled
(`Access-Control-Allow-Origin: *`) tile service, so the canvas stays
exportable — a non-CORS load fails outright rather than tainting. If the
tiles can't load, the figure falls back to a plain background and says
so on the map.

The data layers themselves are drawn using **each service's own
published `drawingInfo` renderer** — the literal colors/classes the
state or USDOT chose, read straight from the layer's REST metadata — so
the symbology is authoritative, not an invented color scheme. One
exception, disclosed in the figure's citation footer: INDOT publishes a
single-symbol renderer (every class the same color), so Indiana's
classes are colored with the standard FHWA palette instead
(byte-identical to the colors MDOT publishes). Every figure includes a
sectioned legend (only the classes actually present in the frame), a
scale bar, a north arrow, a marker for the site, and citations (source
layer names, REST URLs, basemap credit, retrieval timestamp, frame
width) baked directly into the image.

The data layers are deliberately **not** screenshots of a live map.
Only MDOT's service is a classic ArcGIS Server with a
`/MapServer/export` + `/legend` operation; INDOT, WisDOT, and the
nationwide ACUB layer are AGOL-hosted "Query"-only feature services with
no export/legend endpoint at all (confirmed live 2026-07-03). Querying
geometry directly and drawing it with the layer's own renderer works
uniformly across all four sources.

The report only fetches geometry when you click the button (classification
itself never requests geometry, to keep live typing fast) — expect a few
extra round trips in the browser's network log per site. At very wide
frames a service can hit its per-query record limit; the figure then notes
that some segments were not drawn.

## Privacy model (the point of the design)

- **No back end.** The page is plain HTML/JS; the visitor's browser
  queries the public MDOT / INDOT / WisDOT / NTAD / TIGER services
  directly — the identical network path the Excel tool already uses from
  an inspector's laptop. The site operator never sees a coordinate.
- **No damage data.** Input is name + lat/lon — or, since 2026-09-15, a
  road name or street address. WO/DI, applicants, descriptions, categories
  stay in the Excel workbook.
- **Addresses leave the page only on an explicit click.** A typed address
  is sent to the Census Bureau's geocoder (a federal service, over https)
  when — and only when — the `Geocode N addresses` button is pressed; road
  names go to Census TIGERweb automatically, like the Find box. The
  geocoder's JSONP reply executes inside a sandboxed, throwaway iframe that
  cannot reach this page (see "Addresses and road names").
- **Transparency affordances**: a network log at the bottom of the page
  lists every request the page makes; Leaflet is vendored locally
  (`vendor/leaflet/`) so there are no CDN calls; exports (CSV /
  copy-for-Excel) are generated in the browser.
- Residual disclosure: basemap tiles are fetched from Esri for whatever
  area the interactive map shows (streets basemap by default, satellite
  imagery via the layer switcher) and for each PDF report figure's
  frame, and the GIS servers see the queried coordinates in their own
  logs (true of the Excel tool too). Both facts are disclosed in the
  page footer, and the PDF tile fetches appear in the network log.

## Run it

Open `web/index.html` from disk, or serve the `web/` folder from any
static host (GitHub Pages, SharePoint, a local file share). There is no
build step.

## Verify it

```
node build/verify-web-core.mjs
```

Executes the page's `<script id="rr-core">` block (the shipped code, not
a copy) headlessly against the live services, asserting the confirmed
test coordinates from CLAUDE.md §4.2 / §4.2a / §4.2b — including the two
Wisconsin regression points from the increment-4 verdict audit — plus
offline unit checks for coordinate parsing, state auto-detection, and
verdict edge cases. MDOT occasionally returns transient 503s (same
flakiness the Excel tool's re-run-failed button exists for); the page
doesn't cache failures and offers a per-row retry link.

To verify the PDF report feature specifically (a Playwright-driven check
of the actual button click, since it needs a real browser + canvas):

```
cd build/web-tests && npm install && node verify-pdf-report.mjs
```

And the review UX + FIRMette ZIP (map labels, click-to-zoom, Prev/Next,
on-map source layers + legend, sources.html, and a ZIP download validated
end-to-end with Python's zipfile including CRCs):

```
cd build/web-tests && npm install && node verify-review-ui.mjs
```

See each script's header comment for why they stub the network with real
captured fixtures rather than hitting the live services directly — the
query shapes they stub were independently confirmed live via curl first.
