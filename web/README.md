# FHWA Road Checker (provisional title) — web prototype

A single static page that performs RoadReviewer's Workflow 1 (Classify
Roads) in the browser. Paste coordinates into the textarea and each point
is parsed, classified, and pinned on the map automatically — no submit
button, no server, nothing stored. The page is a direct JavaScript port of
`src/modClassify.bas` + `src/modConstants.bas` + `src/modHttp.bas`'s
distance math: same layer URLs, same retired-segment filters, same
exact-intersect-then-buffer fallback (radius adjustable in the UI,
default 250 ft, matching the Excel Search buffer), same per-road
distances, and the same PR #24 verdict model — the **closest** road
segment decides red vs green, yellow only downgrades green with an
explicit Review Reason ("Second road close" / "Nearby FHWA road" /
"Urban boundary edge"), and red never downgrades. Same red/green/yellow
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
tagged **closest** — that's the segment that decided red vs green. Where
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
  by the vendored protomaps-leaflet `light` theme. Browsing the road map
  fetches nothing from a tile CDN. **Satellite stays live Esri**,
  fetched only when switched to; the live Esri street layer remains as
  the automatic fallback when the basemap file is missing (or on
  `file://`). Build pipeline + gotchas: `build/tiles/README.md`.
- **Compact result rows.** Each row shows one line — site name, coords,
  state, verdict badge — and clicking it both selects the site on the
  map and expands the detail (chips, roads, urban area, links).
  **⧉ Expand table** pops the full-detail table for every site at once
  over the map (verdict-tinted, with links), closable via ✕ / Esc /
  backdrop.
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
  stays fully offline; labels appear z14+ (street names), z17+ (house
  numbers). **No POI markers** — the first cut drew POI dots + names
  and the user rejected the clutter. Best-effort: blocked or down, the
  map just shows no detail layer.

## Two input tabs: Coordinate Input + Search & Collect (2026-09-14)

Per user direction the left pane's input area is two tabs:

- **Coordinate Input** — the paste-a-batch flow: the Search **buffer**
  select (renamed from "Search radius") and the coordinates box. The
  State dropdown is GONE — the state is always auto-detected per point.
- **Search & Collect** — the field-collection flow: the Find box
  (state / county / township / city / road via Census TIGERweb — click
  a suggestion and the map zooms there), then an **add-a-point form**:
  Site name, GPS coordinates, and a **Note**. Each added point drops
  into the same shared results table, classifies exactly like a pasted
  point, pins on the map, and **persists in localStorage** until
  removed (per-row "remove" link, or "clear collected"). The note
  shows in the row detail and the pop-out table.
- **Exports**: CSV/Copy-for-Excel/GeoJSON gained a **Note** column, and
  a new **KMZ** export (zipped KML via the FIRMette bundle's store-zip
  builder) writes red/green/yellow pushpins by verdict with name,
  status, class, urban area, roads and note in each placemark — the
  same conventions as the Excel tool's KML.

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

Site pins keep taking the row's verdict color (red / green / yellow, the
same buckets as the row tint and KML pushpins) once a pasted point
classifies; they got a bolder white ring and slightly larger radius so
the verdict reads clearly over the class-colored live road lines.

The results list gets a matching **row filter**: a text box that hides
result cards whose text doesn't match (site name, road/street names,
verdict, class, urban area), and an **"in map view"** checkbox that
keeps only sites inside the current map bounds — Find a county or
township first and the list filters to it, live as you pan. Both are
purely visual: exports, Prev/Next stepping, and classification always
cover every site.

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
- **No damage data.** Input is name + lat/lon only. WO/DI, applicants,
  descriptions, categories stay in the Excel workbook.
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
