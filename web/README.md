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

## PDF report

Click **Download PDF Report** to generate a PDF with a cover page (a
summary table of every classified site) followed by one page per site.
Each site page shows two figures:

1. **The state's road functional-class layer** (MDOT/INDOT/WisDOT —
   whichever one produced the verdict), and
2. **The 2020 Adjusted Census Urban Boundary (ACUB) layer**.

Each figure is drawn from a fresh, geometry-including query against that
same live service, using **that service's own published `drawingInfo`
renderer** — the literal colors/classes the state or USDOT chose, read
straight from the layer's REST metadata — so the figure's symbology is
authoritative, not an invented color scheme. Every figure includes a
legend (only the classes actually present near the point), a scale bar,
a north arrow, a marker for the site, and a citation (source layer name,
REST URL, and retrieval timestamp) baked directly into the image.

This deliberately does **not** screenshot a live map. Only MDOT's
service is a classic ArcGIS Server with a `/MapServer/export` + `/legend`
operation; INDOT, WisDOT, and the nationwide ACUB layer are AGOL-hosted
"Query"-only feature services with no export/legend endpoint at all
(confirmed live 2026-07-03). Querying geometry directly and drawing it
with the layer's own renderer works uniformly across all four sources,
and avoids the canvas-taint risk of compositing remote basemap tiles.

The report only fetches geometry when you click the button (classification
itself never requests geometry, to keep live typing fast) — expect one
extra round trip in the browser's network log per figure per site.

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
  area the map shows, and the GIS servers see the queried coordinates in
  their own logs (true of the Excel tool too). Both facts are disclosed
  in the page footer.

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

See that script's header comment for why it stubs the network with real
captured fixtures rather than hitting the live services directly — the
query shapes it stubs were independently confirmed live via curl first.
