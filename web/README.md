# FHWA Road Checker (provisional title) — web prototype

A single static page that performs RoadReviewer's Workflow 1 (Classify
Roads) in the browser. Paste coordinates into the textarea and each point
is parsed, classified, and pinned on the map automatically — no submit
button, no server, nothing stored. The page is a direct JavaScript port of
`src/modClassify.bas` + `src/modConstants.bas`: same layer URLs, same
retired-segment filters, same 150/200-ft fallback buffers, same
`FederalAidVerdict` table, same red/green/yellow buckets. A "Download PDF
Report" button turns the classified points into a citeable PDF — see
below.

## Site-by-site review

Pasted site names are surfaced as labels on the map pins (permanent up to
25 sites, hover past that). Clicking a table row or a pin zooms straight
to that site; **Prev / Next** above the map step through the sites one by
one (wrapping at the ends). While a site is selected and "Source layers"
is checked, the page fetches the authoritative geometry around it — the
state's road-class layer(s) and the 2020 ACUB urban-boundary polygon,
via the same frame-envelope queries the PDF figures use — and draws it
directly on the interactive map in each source's own published renderer
colors. An on-map legend names the exact layers drawn, lists only the
classes present, links each layer to a live ArcGIS view centered on the
site, and links to the citations page below. Per-site layer fetches are
cached, so stepping back and forth doesn't re-query.

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
your point. Wisconsin and the unwired states have no official public app,
so they get only the pinned link. Plus a Google Maps link. (Layer on/off
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
