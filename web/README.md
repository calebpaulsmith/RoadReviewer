# FHWA Road Checker (provisional title) — web prototype

A single static page that performs RoadReviewer's Workflow 1 (Classify
Roads) in the browser. Paste coordinates into the textarea and each point
is parsed, classified, and pinned on the map automatically — no submit
button, no server, nothing stored. The page is a direct JavaScript port of
`src/modClassify.bas` + `src/modConstants.bas`: same layer URLs, same
retired-segment filters, same 150/200-ft fallback buffers, same
`FederalAidVerdict` table, same red/green/yellow buckets.

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
