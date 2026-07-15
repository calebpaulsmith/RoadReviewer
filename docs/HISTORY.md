# RoadReviewer — historical narrative (moved out of CLAUDE.md)

Verbatim sections moved here 2026-07-15 to slim CLAUDE.md (it rides along in
every AI-assistant request). Nothing here governs the CURRENT build — column
orders, sheet layouts and button surfaces described below have been
superseded repeatedly (the authoritative layout is always `src/modConstants.bas`
+ CLAUDE.md §7d-§7f). Read this file for archaeology: why a decision was
made, what a prototype did, or how an increment unfolded.

---

## 2. Existing prototypes — what they do

### 2.1 `GPS Checker - TN updated 3.5.2026.xlsm`

Sheets: `Start Here`, `Data`, `TEMPLATE` (hidden).

Data sheet columns:

| Col | Field |
|-----|-------|
| A | Name |
| B | Latitude |
| C | Longitude |
| D | Description |
| E | Google Maps (formula `HYPERLINK("https://www.google.com/maps?q=" & B & "," & C, "Open")`) |
| F | Bing Maps |
| G | FEMA Map Viewer |
| H | Flood Map Search |
| I | Download Firmette (link only) |
| J | Firmette Status (populated by macro) |
| K | Actions |
| L | View in TDOT Dashboard |
| M | Functional Classification (populated by macro) |
| N | Urban/Rural |
| O | Road Name |
| P | TDOT Status |

VBA modules:

- **Module1.ExportPointsToKML** — prompts for folder + filename, writes a
  KML file of all points on the active sheet, then opens it.
- **Module2** — older/duplicate FIRMette helpers (not invoked from a button).
- **FirmetteDownloader.bas**
  - `DownloadFirmettes` — prompts for WO, DI, Disaster, output folder, then
    for every row hits the FEMA Print FIRMette GP service
    (`https://msc.fema.gov/arcgis/rest/services/NFHL_Print/MSCPrintB/GPServer/PrintFIRMette`),
    polls `submitJob` → `jobStatus` → `OutputFile`, downloads the PDF to
    `WO{wo} DI{di} - {disaster} - {siteName} FIRMette.pdf`, writes status
    back to column J.
  - `LookupFunctionalClassificationTDOT` — queries TDOT's `Road_Segment`
    FeatureServer (`https://services2.arcgis.com/nf3p7v7Zy4fTOh6M/arcgis/rest/services/Road_Segment/FeatureServer/0/query`).
    Strategy: exact point intersect first; if no hit, fall back to a
    150-foot distance search. Returns up to 5 segments, concatenated with
    ` | `. Cells M:O turn red if any matched class is **not** Rural Local,
    Urban Local, or Rural Minor Collector.
  - `CreateCleanDataSheet` — clones the hidden `TEMPLATE` sheet into a new
    `Data1`/`Data2`/… sheet with headers, hyperlink formulas auto-filled
    down to row 1000, lat/lon data validation, frozen header row, and
    autofilter.

Fields queried from TDOT: `FUNC_CLASS`, `RTE_NME`, `OBJECTID`. The Urban/Rural
designation is *embedded inside* the FUNC_CLASS string (e.g. "Urban Local",
"Rural Minor Collector"), so the macro just looks at the prefix.

### 2.2 `Site Inspector Tool 1.xlsm`

Sheets: `Start Here`, `Setup`, `Sites`, `TEMPLATE` (hidden), `MapPages`.

`Setup` sheet stores job-wide values via named ranges:

- `JobWO` (`B3`), `JobDI` (`B4`), `JobDisaster` (`B5`), `JobApplicant` (`B6`),
  `JobMode` (`B8`), `JobOutputFolder` (`B10`).
- `JobMode` is either "Same for all sites" or "Per-site (granular)".

`Sites` sheet columns:

| Col | Field |
|-----|-------|
| A | WO # *(hidden unless granular mode)* |
| B | DI # *(hidden unless granular mode)* |
| C | Site # |
| D | Site Name |
| E | Latitude |
| F | Longitude |
| G | Category |
| H | Description |
| J | Textbox Preview |
| L–Q | Verification hyperlinks (Google, Bing, FEMA Viewer, Flood Map, Firmette download, Google Earth) — hidden by default; revealed by `ShowVerificationLinks` |
| S | FIRMette Status |
| T | Map Status |
| U | FHWA Class |
| V | Urban/Rural |
| W | Road Name |
| X | FHWA Status |

`MapPages` is generated dynamically by `PrepareMapPages` — one landscape
Letter page per site, with a merged cell area where the inspector pastes a
screenshot using **Place in Cell**, plus a small `Textbox_Site_*` shape in
the top-left containing:

```
WO #{wo} DI #{di}
{Applicant}
Site {#} {Site Name}
{lat}, {lon}
Cat {category}, {description}
```

(First line bold, semi-transparent white fill, thin grey border.)

VBA (`SiteInspectorModule.bas`):

- `SelectOutputFolder` — folder picker writes to `B10`.
- `SyncWoDiMode` — hide/unhide WO/DI columns.
- `ShowVerificationLinks` — unhide L:Q.
- `DownloadFirmettes` — same FEMA GP flow as the GPS Checker, but driven by
  Setup sheet values instead of input boxes; writes status to S.
- `PrepareMapPages`, `AddNewPage`, `ExportCombinedMapPdf`,
  `ExportIndividualMapPdfs` — map-page workflow.
- `CopyTextboxToClipboard` — `MSForms.DataObject` (via CLSID, no reference
  needed) puts the textbox string for the active row on the clipboard, so
  it can be pasted into a manually edited textbox.
- `ExportSitesToKML` — writes KML to `%TEMP%` and shells out via
  `cmd /c start ""` so the default handler (Google Earth Pro) opens it.
- `LookupFunctionalClassification` — same TDOT logic as GPS Checker; writes
  to U:X.
- Ported helpers: `DownloadFirmettePdf`, `GetJobStatus`, `HttpGetText`,
  `HttpDownloadPdf`, `JsonGetValue`, `LookupTdotPoint`,
  `QueryTdotRoadSegments`, `ParseTdotFeatures`, `ExtractAttributeBlocks`,
  `ExtractJsonValue`, `JsonUnescape`, `CoordinateToInvariantString`,
  `WaitSeconds`, `CleanFileName`, `ClassToUrbanRural`,
  `ShouldHighlightClassRed`, `JoinStringArray`, `StringArrayCount`.

Network stack: `MSXML2.ServerXMLHTTP.6.0` for GET/JSON,
`ADODB.Stream` (Type 1, binary) for the PDF write, `VBScript.RegExp` for the
hand-rolled JSON parsing.

### 2.3 Findings about the existing design

**Good ideas to keep:**

- Single shared Sites table that every workflow reads from.
- Verification hyperlinks built as formulas off lat/lon — zero macro cost.
- Setup sheet with named ranges for job-wide values.
- FIRMette download pattern (submitJob → poll → OutputFile → write binary).
- Per-row status columns so failures are visible without scrolling logs.
- Red highlight rule for ineligible road classes.
- Textbox overlay on map pages so each PDF has the WO/DI/site stamp baked in.
- KML export → opens in Google Earth Pro by default association.

**Problems / friction points the redesign must address:**

1. **Too many buttons, no clear flow.** The Site Inspector Tool has eight
   public subs, all visible on one sheet. A first-time user does not know
   the order. We need a clearly numbered, opinionated workflow.
2. **Tennessee-only.** Both tools query TDOT's Road_Segment service
   directly. Michigan (V1) needs MDOT's NFC + ACUB layers. The eligibility
   rule logic must be rewritten to handle states where Urban/Rural is *not*
   baked into the functional-class string (MDOT's NFC layer carries class
   and urban-area flags as separate fields — to be confirmed).
3. **Hand-rolled JSON parser via RegExp.** Works, but fragile against
   nested objects, escaped quotes, and reordered fields. Acceptable for
   simple `attributes` blocks, but we need to keep the schema simple and
   add defensive tests.
4. **GPS Checker writes to *active sheet*** (`Set ws = ActiveSheet`),
   which means a user on the wrong tab silently writes garbage into a
   non-data sheet. Site Inspector Tool fixes this by hard-binding to
   `Sites`; we should do the same.
5. **MapPages assumes screenshots are pasted by hand.** This is fine for
   V1 but the inspector still has to flip between MDOT/Google/FEMA, take a
   screenshot, paste in cell. We can streamline this by opening a curated
   set of imagery URLs in one click (workflow 2) and possibly by
   automating a static-basemap fetch (see open question).
6. **`Set ws = Nothing` checks with `On Error GoTo 0`** scattered through
   the code. Style consistency matters when junior staff inherit this.
7. **`Shell "explorer.exe …"`** and `Shell "cmd /c start ""`** — fine on
   typical govt laptops but worth listing as a known dependency in case
   AppLocker policies block it.
8. **No retry / no resume.** If `submitJob` fails for one row, the row is
   marked failed and the run continues. There's no "re-run just the
   failed rows" command.
9. **Inconsistent column letter constants vs. column-index constants** in
   the two files. The redesign should pick one (named constants for column
   indices) and stick with it.
10. **`VBA Stomping was detected`** on the Site Inspector Tool by olevba
    (source code and P-code differ). Usually benign — it means the file
    was saved by a different Office version than it was authored on — but
    a clean rewrite removes the warning.
11. **No address input.** Both tools accept only decimal-degree lat/lon.
    PA inspectors often have street addresses; needs a geocoder (see
    open question).

---

## 7a. Implementation status (V1)

**Increment 2 — every capability built, every automated verifier passes
against the live services.** The fixes from the cloud-only authoring pass
(see commits `fa26093` and `ece6c29` for the gory detail) landed during
the local smoke-test pass; what was a "structurally verified only" build
in increment 1 is now a workbook that drives Workflows 1, 2 and 3 end to
end. End-user (button-click) smoke is still the final gate — the
automated verifiers exercise the same code paths but in headless mode.

**Increment 3 — Indiana + Wisconsin NFC wired, schema-verified but not
yet Excel-tested.** `modClassify.bas` now dispatches per state (MI/IN/WI)
instead of hardcoding MDOT; see §4.2a/§4.2b for the live-verified INDOT
and WisDOT schemas this was built against. The schema/field-name research
was done by fetching each service's `?f=pjson` metadata and running live
point-intersect queries directly (this repo's cloud sandbox can reach
`gisdata.in.gov` and `services5.arcgis.com`, unlike `mdotgis.state.mi.us`
which needed the local-workstation probe workflow in §4.2's dev notes).
What has **not** been exercised yet is the VBA HTTP stack itself
(`MSXML2.ServerXMLHTTP` through `modHttp.HttpGetText`) against these two
new endpoints — that requires the local Windows+Excel rebuild-and-verify
loop in §9.1/§9.2, same as every other capability in this table. Run
`verify-classify.ps1` (now covers all three wired states, §9.2) before
trusting IN/WI output on a real WO.

**Increment 4 — federal-aid verdict audit, three bugs fixed in
`modClassify.bas`.** Prompted by a previously-observed field issue: a site
point sitting just outside an ACUB polygon (e.g. on the wrong side of a
road that itself touches the urban boundary) got flagged Rural instead of
Urban. Root-caused and fixed, plus two more bugs found auditing
`FederalAidVerdict`/`PrefixedClass` against the §4.2 eligibility table
while in there:

1. **ACUB boundary buffer decoupled from the road-search buffer, with a
   200 ft floor (`AcubBufferFeet`/`ACUB_MIN_BUFFER_FEET`).** The ACUB
   point-in-polygon check already fell back to a buffered search when the
   exact point missed (`QueryWithFallback`), which is what makes the edge
   case survivable at all — but it reused `JobBufferFeet`, the same knob
   Setup's help text says to *narrow* "for dense urban grids" to get more
   precise road matching. Narrowing it for that purpose silently
   shrank the urban-boundary tolerance too, reopening the exact bug
   above. Now `QueryWithFallback` takes an explicit fallback distance per
   call site; ACUB always gets `Max(JobBufferFeet, 200 ft)` while road
   queries keep using the raw `JobBufferFeet` as before.
2. **Rural Minor Collector mislabeled "...Rural Local".** `FederalAidVerdict`
   correctly withheld the federal-aid flag for a rural class-6 segment (no
   behavior bug in the flag itself), but the `Else` branch that builds the
   display string unconditionally appended "Local" regardless of which
   class was actually found — so the Federal Aid Status column read
   "Non-federal aid - Rural Local" for a road the FHWA Class column (right
   next to it) correctly labeled "Minor Collector". Never caught because
   no test coordinate exercised a rural class-6 segment. Fixed: a
   dedicated `hasRuralMinorCollector` branch now outputs "Non-federal aid
   - Rural Minor Collector", matching the eligibility table this code was
   always supposed to implement.
3. **Codes 1-3 (Interstate/Freeway/Other Principal Arterial) weren't
   Urban/Rural-prefixed** the way codes 4-6 were, even though the §4.2
   table's "Federal aid - <Urban/Rural class>" format implies every
   federal-aid class gets the prefix. `PrefixedClass` simplified to
   always prefix + reuse `FunctionalSystemLabel`, which was already
   returning byte-identical text for 4-6 (`"Minor Arterial"`, `"Major
   Collector"`, `"Minor Collector"`) — so this is a pure fix, not a
   behavior change, for the classes that were already correct.

Two live-verified WI coordinates were added to `verify-classify.ps1`
(rows 9-10) as regression tests for fixes #2 and #3: STH 52 near
Rhinelander (`45.169879, -89.102452`, FED_FC_CD=6/Rural, confirmed clean
of any other class within 200 ft and outside every ACUB polygon) must
read exactly "Non-federal aid - Rural Minor Collector", and I-94 through
Eau Claire (`44.764850, -91.406533`, FED_FC_CD=1, inside the "Eau Claire,
WI" ACUB) must read exactly "Federal aid - Urban Interstate". Fix #1 (the
ACUB buffer floor) doesn't have an automated regression test - it only
matters when `JobBufferFeet` is manually narrowed below 200 ft, which
none of the verifiers configure.

**Increment 5 — live cell updates during long-running workflows.**
`ClassifyRows`, `FirmetteRunRows` and `GeocodeAddresses` all previously
set `Application.ScreenUpdating = False` for the duration of their
per-row loop, then flipped it back on when done - the button-owning
sheet (e.g. "1. Classify Roads") stayed on screen and the Sites table
only visibly updated all at once at the end. Each row in these three
workflows is dominated by network latency (2-3 HTTP round trips for
classify, a submitJob/poll/download sequence for FIRMettes, one Census
call for geocoding), so the redraw-suppression bought essentially
nothing while hiding the one thing an inspector watching a multi-minute
run most wants to see: progress. All three now `ws.Activate` the Sites
sheet before the loop, drop the `ScreenUpdating = False`, and `.Select`
the row's cell being written each iteration so the selection visibly
tracks down the sheet as it fills in. `PrepareMapPages`/`AddMapPage`
were deliberately left alone - that loop is CPU-bound shape/textbox
creation with no per-row network wait, so suppressing redraws there is
a real performance win, not a UX cost; it already activates MapPages
once the run finishes.

| Capability | Module | Status |
|---|---|---|
| Skeleton (Home/Setup/Sites + 3 workflow sheets, buttons, named ranges) | modBuild | **tested** (§5.2 — verify-skeleton.ps1) |
| Sites hyperlinks + lat/lon validation + INELIGIBLE red highlight | modBuild | **tested** (§5.3 — verify-skeleton.ps1) |
| Workflow 1 — Classify Roads, Michigan (NFC 353 + ACUB + route 543, eligibility, re-run failed, state gate) | modClassify | **tested** (§5.4/§5.7/§5.8 — verify-classify.ps1, verify-rerun-and-state.ps1) |
| Workflow 1 — Classify Roads, Indiana (LRSE_Functional_Class + centerline road name) | modClassify | built against a live-verified schema (§4.2a); needs a local `verify-classify.ps1` run before first real use |
| Workflow 1 — Classify Roads, Wisconsin (state-trunk + local-roads fallback, category-code normalization) | modClassify | built against a live-verified schema (§4.2b); needs a local `verify-classify.ps1` run before first real use |
| Workflow 1 — Classify Roads, Minnesota / Illinois / Ohio (bare FHWA 1-7, Indiana shape; OH route names) | modClassify | **tested** (PR #36 — verify-classify.ps1 all six states + verify-web-core, 2026-07-15, live) |
| Workflow 2 — Review Imagery (open curated set for selected rows) | modImagery | built; uses the same URL templates §5.3 already verified resolve correctly. End-user click test pending |
| Geocode addresses → lat/lon (never overwrites) | modGeocode | built; no automated verifier yet (one-shot Census call) |
| KML export + Sites-table CSV export | modMaps, modExport | built |
| Output-folder resolution (§8.9) | modMaps | built (exercised via verify-firmette-maps.ps1) |
| Workflow 3 — Download FIRMettes / Re-run failed FIRMettes | modMaps | **tested** (verify-firmette-maps.ps1 — 17.9s end-to-end against FEMA Print FIRMette GP, 828 KB PDF written) |
| Workflow 3 — Prepare Map Pages / Export Combined Map PDF | modMaps | **tested** (verify-firmette-maps.ps1 — MapPages sheet + textbox shape created, 50 KB Location Map PDF exported) |

**Bugs found and fixed during increment 2** (full commit-message detail in
`ece6c29`):

1. `Worksheet.DisplayGridlines = False` is a compile error — that property
   lives on `Window`. New `HideGridlines` helper activates the sheet and
   flips `ActiveWindow.DisplayGridlines`.
2. `Public gTracePath As String` was placed between two Subs in modUtil —
   VBA's "Only comments may appear after End Sub" compile error. All
   module-level Public state moved to the top of modUtil.
3. `NewRegex(...)` had a parameter named `global` — that's a VBA reserved
   word (synonym for `Public`). JIT compilation later threw "Sub or
   Function not defined" on the callers and Excel sat in VBE break mode
   forever. Renamed to `isGlobal`.
4. The narrow JSON regex defaulted `IgnoreCase=True`, so `FirstString("NAME")`
   on an ArcGIS response found the lower-case `"name":"OBJECTID"` field-
   metadata entry instead of the upper-case attribute `"NAME":"Kalamazoo, MI"`.
   ArcGIS attribute names are case-sensitive in the JSON anyway —
   defaulted IgnoreCase to False.
5. The repo's `build/` folder carries an `Everyone Deny
   DeleteSubdirectoriesAndFiles` ACE from the Claude Code sandbox, which
   breaks Excel's SaveAs (it writes a temp file then deletes/renames).
   `build/build.ps1` now defaults the OutPath to `%TEMP%`, and a
   `.gitignore` keeps the binary out of the repo either way.

**Headless plumbing** — every public workflow Sub checks `gHeadless` and
suppresses its success/failure MsgBox when an automation host (build.ps1,
verify-*.ps1) sets it via `Application.Run "SetHeadless", True`. Cell +
StatusBar state remain, so results stay observable when running headless.

**Increment 6 — the two-product split (2026-07-05).** The single workbook
became `RoadReviewer.xlsm` (standard) + `Site Inspector Review Tool.xlsm`
(inspector); six sheets collapsed to three; geocoding folded into Check
Roads; new column order + Sites toolbar row. Full design, verification
results and two bug fixes found during the pass in **§7c**. The
per-capability status above still describes the shared logic accurately;
run the §9.2 runbook (build both, verify each product) after any `src/`
change.

Open the built workbooks in Excel for the button-click smoke pass:
imagery one-click (§5.5), the address auto-geocode, and any UI feel-test
before handing off to a non-developer co-worker (§5.10).

---

## 7b. FHWA Road Checker (provisional title) — public web prototype

A separate product exploration living in `web/`: RoadReviewer's Workflow 1
as a public static page, aimed at applicants (county road commissions,
municipal engineers) rather than inspectors. Full rationale and privacy
model in `web/README.md`; the design conversation that produced it is
summarized here so it isn't relitigated:

- **Client-side only, by design.** No back end, no accounts, no storage,
  no analytics. The visitor's browser queries the same public MDOT /
  INDOT / WisDOT / NTAD / TIGER endpoints the Excel tool queries (every
  one of them was confirmed to send CORS headers, 2026-07-02 — including
  `mdotgis.state.mi.us`, which is also reachable from this repo's cloud
  sandbox now, contrary to the older note in §7a). This is the answer to
  "the audience won't upload damage data to the internet": the page never
  asks for damage data (name + lat/lon only, no WO/DI/applicant fields)
  and has no server to upload to. Static hosting also means "thousands of
  visitors" costs nothing; the scaling risk is the upstream state servers,
  mitigated by per-visitor IPs and the small per-point query count.
- **Paste-and-see UX.** Coordinates pasted into the textarea are parsed,
  classified and pinned on a Leaflet map automatically (debounced input
  event — no submit button). Rows/pins use the same red/green/yellow
  federal-aid buckets as the Sites table and KML export. Results render
  as one compact flex card per site rather than a spreadsheet-style
  table (2026-07-05 redesign: no horizontal scroll, thin rows, solid
  verdict badge + tinted row): the Excel tool's pipe-joined class
  string is unpacked into "N road segments within 200 ft:" plus one
  chip per road/class, swatched in the standard FHWA class colors —
  the same palette the map overlay and PDF figures use. WI chips pair
  name + class (both WI layers carry names on the class feature);
  MI/IN class layers have no name field, so their chips are class-only
  with a ×n multiplier, next to the separate Route/TIGER street names.
  rr-core's classifyPoint exposes `segments`/`bufferFt`/`streetList`
  for this, alongside the unchanged flat fields. A visible
  network log lists every request the page makes; Leaflet is vendored
  locally so there are no CDN calls. Exports: CSV download +
  copy-for-Excel TSV, both generated in the browser — these keep the
  original flat pipe-joined columns for spreadsheet hand-off.
- **Same logic as Excel.** `web/index.html`'s `<script id="rr-core">`
  block is a hand-port of `modClassify.bas`/`modConstants.bas` (same
  where-clauses, fallback buffers incl. the ACUB floor, verdict
  table, state gate). The two must be kept in sync by hand — there is no
  shared source between VBA and JS. `build/verify-web-core.mjs` executes
  that exact script block headless against the live services and passed
  on all §4.2/§4.2a/§4.2b test coordinates (2026-07-02), including the
  two increment-4 WI regression points; a Playwright DOM smoke (paste →
  table rows + colored markers, stubbed services) also passed. MDOT
  throws occasional transient 503s; failed rows aren't cached and get a
  per-row retry link (web analog of F12).
- **PR #24 verdict model ported to the web (2026-07-08).** rr-core now
  mirrors the Excel coloring overhaul end to end: per-road distances
  (geometry + point-to-segment math, port of `modHttp.MinDistanceFt`),
  three-state ACUB (exact / boundary-edge / rural, 250-ft floor),
  closest-road-drives-red/green, yellow-only-downgrades-green with a
  `reviewReason` field (Second road close / Nearby FHWA road / Urban
  boundary edge), and the 250-ft default buffer. Same pass added: a
  **Search radius** select (50–1,000 ft; part of the result cache key),
  distance-annotated class chips with the closest one tagged (it decides
  the color), a merged nearest-first "Roads:" list (= the Excel Road
  Name column), a "How the colors are decided" explainer on the page +
  a step-by-step `sources.html#verdict` section, per-row **ArcGIS map**
  link (parity with the Excel AGOL NFC Layer column's
  `URL_NFC_MAPVIEW*` templates) and **Public map** link (official state
  app root, all six states), and a **Download GeoJSON** export matching
  `modMaps.WriteSitesGeoJson`'s property shape for drag-drop onto AGOL
  maps. CSV/TSV exports gained Review Reason + the distance-annotated
  road list. All three web verifiers were updated and pass (42/42 core
  checks incl. every live §4.2* coordinate; the review-UI test now also
  exercises the yellow "Second road close" path, the radius control,
  the new links, and the GeoJSON download).
- **PDF report.** A "Download PDF Report" button (jsPDF, vendored like
  Leaflet) produces a cover page + one page per classified site. Each
  site page has **one combined, page-filling map**: an Esri World
  Street Map tile basemap (roads + street names for orientation) at the
  bottom, the ACUB polygon tinted over it, and the state's
  functional-class polylines on top (WI draws both the state-trunk and
  local-roads layers), queried by frame **envelope**
  (`esriGeometryEnvelope`, live-verified against all four services
  2026-07-05: 51-823 features per 0.75 mi urban frame, no record-limit
  hits) so the figure shows the surrounding network for context, not
  just the verdict segment. The figure is projected in Web Mercator so
  the tiles composite pixel-perfect; the interactive Leaflet map
  likewise defaults to the same streets basemap, with a layer-switcher
  back to satellite imagery. Frame width comes from a
  "PDF map width" select next to the button (300/600/1200/2400 m
  half-width; default 600 m ≈ 0.75 mi across - deliberately further out
  than the original 250 m figure; retune by editing the option values).
  Below the figure each page carries clickable live source links (jsPDF
  `textWithLink`); the review legend carries the same links. These come in
  two tiers (reworked across PR #17 → PR #18, 2026-07-06):
  **(1) primary reference** — each state's own official public app
  (MI = MDOT's "NFC, NHS & ACUB" ArcGIS Experience; IN = INDOT's
  "Functional Classification & Urban Area Boundary" Experience), shown
  first and, in the PDF, bold-starred. Both show functional class *and*
  the adjusted urban boundary by default, so the one reference answers
  both questions. **(2) first-tier pinned links** — the FEMA-hosted
  ArcGIS Map Viewer centered+markered on the exact site (MI = curated
  NFC/ACUB `webmap=6a1702b9…`; IN/WI/ACUB = FeatureServer `url=`
  side-loads), which actually put the user on their point.
  **History / gotcha:** PR #17 first tried to make the official Experience
  app the *only* link, deep-linked with the Experience Builder
  hash-parameter form `#<widgetId>=center:lon,lat,…` (widget IDs
  MI `widget_167`, IN `widget_6`, recovered from each app's config at
  `sharing/rest/content/items/<id>/data?f=json`). In a real browser this
  **mis-navigated** ("went to some random site") — EXB's router balks at
  the widget hash combined with a stale `/page/<name>` path segment. PR
  #18 rolled that back: the official app is linked at its **canonical root
  URL** (the `url` field from the AGOL item — no `/page/…`, no hash) as
  the reference, and the proven FEMA-viewer pin links are restored as the
  first-tier clickable links. Experience Builder also can't toggle layer
  visibility by URL, but the wanted layers are the app defaults, so that's
  moot. WisDOT publishes **no** statewide public functional-class app
  (static county/urban PDFs + login-gated WISLR only), so Wisconsin and
  the unwired states get only the pinned FEMA-viewer link, no official
  reference. Plus a Google Maps link. See `STATE_APP` / `officialAppLink`
  / `miWebmapLink` / `liveLayerLink` in `web/index.html` and the per-state
  notes in `sources.html`; the Excel tool's `URL_NFC_MAPVIEW*` constants
  (§9.3b) are unchanged and still use the older webmap/side-load patterns.
  Rather than
  screenshotting a live map, the figure is drawn on a plain `<canvas>`
  styled with **each layer's own published `drawingInfo.renderer`** (the
  state's/USDOT's actual class-to-color mapping, read from the layer's
  REST metadata) - so the symbology is authoritative rather than
  invented. This choice was forced by a live probe (2026-07-03): only
  MDOT's service is a classic ArcGIS Server exposing `/MapServer/export`
  + `/legend`; INDOT, WisDOT, and the nationwide ACUB layer are
  AGOL-hosted feature services with `"capabilities": "Query"` only - no
  export/legend operation exists for them at all, confirmed by both a
  direct `/MapServer/export` 400/404 and the FeatureServer root's
  `capabilities` field. Drawing every source's own renderer client-side
  sidesteps that inconsistency uniformly. The street-tile basemap
  underneath doesn't reintroduce canvas taint: tiles load with
  `crossOrigin="anonymous"` against Esri's CORS-enabled
  (`Access-Control-Allow-Origin: *`, confirmed live 2026-07-05) tile
  service, which errors on a non-CORS response instead of tainting, and
  the figure falls back to a plain background with an on-map note if
  tiles don't load. Vector paths/rings from queried geometry, a marker,
  a scale bar, a north arrow, and a sectioned legend/citations
  (including the basemap credit) are drawn onto the same canvas.
  Renderer traps found
  during this work: Wisconsin's state-trunk layer keys its renderer off
  `FC_CD` (WisDOT's own code), not the `FED_FC_CD` field the classifier
  itself uses - the report's query fetches both fields, using `FC_CD`
  only for color-matching; INDOT's layer publishes a single-symbol
  renderer (every class one color, confirmed live 2026-07-05), so
  Indiana substitutes the standard FHWA palette (byte-identical to
  MDOT's published colors) via `fhwaFallbackDrawingInfo`, disclosed in
  the citation footer; and the jsPDF doc must be created with
  `compress: true` or each figure embeds as ~5 MB of raw RGBA instead of
  ~90 KB total. Verified end-to-end with a real jsPDF download
  in a real (headless) browser via `build/web-tests/verify-pdf-report.mjs`
  (which now inflates the compressed PDF streams before grepping for
  text, and also asserts the live-source link annotations), stubbed with
  response fixtures captured live from MDOT/NTAD (see that file for why:
  this sandbox's Chromium couldn't complete TLS through the outbound
  proxy that curl/Node's own fetch use fine, so the envelope query
  shapes were confirmed correct via curl first and the browser-side
  drawing/PDF-assembly code was then verified against those real
  captured responses) - a plain user's browser with normal internet
  access isn't expected to hit that proxy limitation.
- **Site-by-site review.** Pasted names appear as labels on the map pins
  (permanent ≤25 sites, hover beyond); clicking a row/pin zooms to that
  site (z17); Prev/Next buttons step through sites with wrap-around.
  While a site is selected, a "Source layers" toggle (default on) draws
  the authoritative geometry around it straight onto the Leaflet map —
  the same frame-envelope fetchers and published-renderer symbology the
  PDF figures use (`fetchClassLayers`/`fetchAcubLayer`, 400 m half-width
  frame, results cached per site) — plus an on-map legend naming the
  exact layers, listing only the classes present, and linking each layer
  to a live ArcGIS view and to sources.html. The map is created with
  `zoomAnimation: false`: Leaflet's `_tryAnimatedZoom` returns true
  without moving when a zoom animation is already in flight, so clicking
  Next during the ~250 ms animation would randomly not zoom (and stuck
  `_animatingZoom` breaks headless verification outright — root-caused
  in-session with a Leaflet-internals trace).
- **sources.html.** Per-state citations page: organization, service URL,
  exact layer names, fields read, and every schema quirk documented in
  §4.2/§4.2a/§4.2b (retire-date/record_status filters, WI category
  codes, INDOT's single-symbol renderer + FHWA-palette substitution,
  ACUB buffer floor, MDOT trunkline-only names). Linked from the page
  header, each result row's "Source" link (anchored to the row's state),
  and the review legend.
- **FIRMette batch → ZIP.** "Download FIRMettes (ZIP)" drives FEMA's
  Print FIRMette GP service per site from the browser (submitJob → poll
  2 s/max 90 → OutputFile → PDF; every step's CORS confirmed live
  2026-07-03, including the output PDF), 2 jobs concurrent, capped at
  20 sites per run (FEMA renders each ~1 MB PDF fresh). The ZIP is
  assembled in-page by a dependency-free STORE-only writer (PDFs are
  already compressed); per-site failures are reported and don't sink
  the batch. `build/web-tests/verify-review-ui.mjs` covers all of the
  above — the ZIP is validated with Python's zipfile including CRCs.
- **Open items:** the "FHWA Road Checker" name risks implying agency
  affiliation (page carries an "unofficial, not affiliated" disclaimer;
  consider "Federal-Aid Road Checker"); basemap tiles necessarily reveal
  the viewed area to Esri (disclosed in the footer); state auto-detect
  uses rough bounding boxes (wrong guesses fail soft to "review manually"
  + dropdown override). The MI/IN official-app links now open the app's
  canonical root (default statewide extent) — the fragile Experience
  Builder coordinate deep-link was removed after it mis-navigated in a
  real browser (PR #18), so the exact-location duty sits on the FEMA-viewer
  pin links, which resolve pin+zoom even if a side-loaded overlay doesn't
  render.
- **Hosting.** `.github/workflows/pages.yml` deploys `web/` (no build
  step — the folder is uploaded as-is) to GitHub Pages on every push to
  `main` that touches `web/**`. This only takes effect once the repo's
  Settings > Pages > Build and deployment > Source is set to
  "GitHub Actions" — a one-time manual toggle; no available tool can
  flip a repo setting like that from this session.
  **Re-run gotcha:** every clean (attempt-1) run deploys fine, but
  *re-running* a run fails with "Multiple artifacts named github-pages …
  Artifact count is 2" — attempt 2's `upload-pages-artifact` adds a second
  artifact next to attempt 1's, and `deploy-pages` then sees two. The
  workflow is split into `build` + `deploy` jobs so "Re-run failed jobs"
  (deploy only) is safe; to recover a wedged deploy, trigger a **fresh**
  run (push, or Actions > Run workflow / `workflow_dispatch`) rather than
  "Re-run all jobs".

---

## 7c. Two-product split + workflow simplification (2026-07-05)

The single six-sheet workbook was split into two products for two distinct
audiences, per user direction. All decisions below were made explicitly in
that design session — do not relitigate.

### The two products

Both are built from the same `src/` tree by `build\build.ps1` and share the
same three-sheet shape — **Start Here** (all inputs + every action button),
**Sites**, **Sources** (per-state citations + quirks, `modSources.bas`).
There are no navigation-only buttons and no per-workflow sheets anymore.

| | RoadReviewer.xlsm (Standard) | Site Inspector Review Tool.xlsm (Inspector) |
|---|---|---|
| Audience | PDMGs, state/local partners, project reviewers | FEMA PA site inspectors |
| Inputs | State, Output Folder (optional), AGOL URL (optional) | + WO, DI, Disaster, Applicant, Search buffer |
| Buttons | Check Roads, Re-run Failed, Photo Links, CSV, KML, Send-to-AGOL, Build/Reset | + Download/Re-run FIRMettes, Prepare Map Pages, Add Blank Map Page, Export Combined Map PDF |
| Sites columns | WO#, DI#, FIRMette Status, Map Status **hidden** (and dropped from its CSV) | all visible |
| Output folder default (blank) | same folder as the .xlsm (falls back to the §8.9 probe if unsaved / an https path) | §8.9 OneDrive-FEMA pattern, unchanged |

### Product identity plumbing

The product id is baked in at build time as a **hidden defined name**
`RR_Product` (`modUtil.SetProduct`, called by build.ps1 via
`Application.Run` before `BuildWorkbookSafe`). Runtime code branches via
`ProductIsInspector()` / `ProductTitle()`; a missing name defaults to
Inspector (the superset) so pre-split workbooks keep full behavior. The
in-Excel "Build / Reset Workbook" button rebuilds the same product.
`COL_*` constants are shared — the standard product *hides* inspector-only
columns rather than having its own column map, which is what keeps
modClassify/modMaps/modExport product-agnostic.

### Workflow simplifications (both products)

1. **One primary action.** "Check Roads" = auto-geocode (any row with an
   Address but no coords, Census geocoder, never overwrites typed coords)
   then classify. The standalone Geocode button/sub is gone
   (`modGeocode.GeocodeRow` is a helper called per-row from
   `modClassify.ClassifyOneRow`). Geocode failures write
   `Failed - geocode: …` into Federal Aid Status so **Re-run Failed Rows**
   (renamed from ReRunFailedClassifications; CheckRoads replaced
   ClassifyAllRows) retries them too.
2. **Sites toolbar row.** Sites row 1 now holds a hint line + two
   free-floating buttons (Check Roads, Photo Links for selected rows) so
   the paste → classify → review loop never leaves the sheet. Header moved
   to **row 2**, first data row is **3** (`SITES_TOOLBAR_ROW` /
   `SITES_HEADER_ROW` / `SITES_FIRST_DATA_ROW`).
3. **Paste-friendly column order.** Latitude | Longitude | Description are
   contiguous; Address sits to the RIGHT of Description (rarely used).
   New canonical order: WO(1) DI(2) Site#(3) Site Name(4) Lat(5) Lon(6)
   Description(7) Address(8) Category(9) Costs(10) Work Completion(11)
   Geocode Status(12) | links: Google Maps(13) Street View(14) Bing(15)
   **Google Earth(16, new)** FEMA Viewer(17) FIRMette Portal(18) NFC
   Map(19) | results: FHWA Class(20) Urban/Rural(21) ACUB Name(22) Road
   Name(23) Street Name(24) Federal Aid Status(25) | FIRMette Status(26)
   Map Status(27) AGOL Map(28).
4. **Input/output tinting.** Input columns are light yellow with
   "(optional)" header suffixes where applicable; lookup/result columns
   light grey; the tri-color verdict conditional formats win over the
   static tints when they match.
5. **Google Earth link** (`URL_GEARTH`, earth.google.com/web/search) added
   as a column and to the multi-open photo set (now 5 tabs/site) — it was
   the by-name-requested pre-disaster imagery source.
6. **Search buffer hidden in standard.** No cell/named range is created;
   `SetupValue` returns "" for the missing name and `BufferFeet()` falls
   back to the 200 ft default. Inspector keeps the field.

### Verified 2026-07-05 (all on the live services, this machine)

verify-skeleton (both products), verify-classify (MI/IN/WI + the
auto-geocode row, all pass), verify-rerun-and-state, verify-firmette-maps
(828 KB FIRMette + Location Map PDF), verify-blank-wodi, plus a manual
standard-product CSV test (lands next to the workbook, drops the four
inspector-only columns, includes Google Earth). Two fixes came out of that
pass: (a) the §4.2a Indiana test point 39.9876,-86.0128 was mislabeled
"rural Hancock County" — it is inside the Indianapolis ACUB (test now
expects Urban Local); (b) a CSV off-by-one where keying the comma off
"line still empty" swallowed leading empty fields (blank Site #) — the
separator is now counted per EMITTED column (`modExport.CsvLine`).

### Migration note

Workbooks filled in under the old single-product layout do NOT migrate:
the column order changed and the header moved to row 2. Old files keep
working as-is (their embedded VBA is self-consistent); new work should
start from the new deliverables.

### Front-page disclaimer, boundary handling, citation page, version stamp (PR #21)

Four additions after the split, per user direction:

1. **Front-page disclaimer** (`modBuild.DisclaimerBlock`) — a red-bordered
   box near the top of Start Here on **both** products: the tool does NOT
   authoritatively identify FHWA roads, it flags high-probability
   candidates for human review, is not authoritative, every coordinate
   must be verified by a human on the source map, and results do not
   constitute a federal-aid / funding / eligibility determination. Same
   wording echoed on the Sources sheet ("Read this first") and in
   `web/index.html`. Adding it reflowed both Start Here layouts (buttons
   are positioned by `ws.Rows(N).Top`, so every row index below the box
   shifted down; named-range cells moved accordingly — the skeleton
   verifier checks `RefersTo` presence, not fixed rows, so it still
   passes).
2. **Boundary-road handling is intentional and now documented.** The
   answer to "does a point on/near an urban boundary get counted Urban?"
   is **yes** — the ACUB check's fallback radius has a 200 ft floor
   (`ACUB_MIN_BUFFER_FEET`, §7a increment 4), so a point on a
   boundary-forming road or a few feet onto the rural side resolves
   **Urban**, biasing toward a federal-aid flag for review rather than
   silently dropping a boundary road. This is now spelled out on the
   Sources sheet ("BOUNDARY ROADS") and in the inspector's Search-buffer
   note, not just in this design doc.
3. **Citation page** = the `Sources` sheet (`modSources.bas`), already
   added in the split; PR #19 enriched it with the disclaimer echo, the
   boundary caveat, and a verified-on/version footer. It mirrors
   `web/sources.html` (org, service URL, exact layer, fields read, every
   schema quirk per state).
4. **Version/PR stamp** — `BUILD_REFERENCE` constant in `modConstants`
   (`"PR #21"`), shown as a small grey label at the bottom of Start Here
   (`modBuild.VersionLabel`) and in the Sources footer, so a shared copy
   is traceable to the build/PR it came from. Bump it each release.
   `verify-skeleton.ps1` asserts the disclaimer text, the eligibility
   clause, the `PR #` label, and the `BOUNDARY ROADS` caveat.

### State-selector labels, output-folder display, two-section Sources (PR #22)

1. **State dropdown labels the unwired states.** `STATE_LIST` is now
   `"WI,IN,MI,MN (not wired),IL (not wired),OH (not wired)"` so a user sees
   at a glance which states classify roads. `modUtil.BareStateCode()` strips
   the `" (not wired)"` suffix back to the bare 2-letter code; `ClassifyRows`
   and `modExport.NfcMapUrlForRow` both call it instead of `UCase$` on the
   raw cell. (It is named `BareStateCode`, not `StateCode`, on purpose: the
   callers' local variable is `stateCode`, and a same-named function is
   shadowed by the case-insensitive local, compiling `stateCode =
   StateCode(...)` as an "Expected array" error — the NfcWired/nfcWired
   trap. §9.3.) The wired states stay bare (`"MI"`/`"IN"`/`"WI"`), so
   `SetNfcMapFormula`'s Excel-side `=IF(state="MI",…)` comparisons are
   unchanged. `verify-rerun-and-state.ps1` now sets `"MN (not wired)"` to
   exercise the normalizer end-to-end.
2. **Standard product's Output Folder shows its own directory.** The two
   explanatory notes under the inputs were deleted; instead the Output
   Folder cell carries a live `CELL("filename")` formula
   (`modBuild.SetOutputFolderDefault`) that displays the workbook's folder
   (or blanks for an unsaved / `http` SharePoint path — the exact case
   `ResolveOutputFolder` already falls back on, so behavior is unchanged).
   Standard only; the inspector keeps its blank cell + OneDrive-FEMA
   default. Browse/typing overwrites the formula.
3. **Sources sheet split into two sections** (`modSources.bas`):
   **(1) SOURCES** — per state, the official public map + the data service,
   as short bluebook citations ending "available at: <url>" (whole line
   hyperlinked); **(2) QUIRKS & CAVEATS** — the same schema quirks in plain
   language. Official public functional-class maps + reference REST layers
   for all six states were verified live 2026-07-05:

   | State | Official public map | Class REST layer |
   |---|---|---|
   | MI | MDOT "NFC, NHS & ACUB" Experience | mdotgis…/FeatureServer/353 (wired) |
   | IN | INDOT "Functional Class Map" Experience | gisdata.in.gov…/FeatureServer/22 (wired) |
   | WI | WisDOT function.aspx (static county/urban PDFs; **no statewide interactive map**) | services5.arcgis.com FFCL_gdb/FeatureServer/3 (wired) |
   | MN | MnDOT EMMA (`webgis.dot.state.mn.us/emma/`) | dotapp9…/mndot_commonlayers2/MapServer/11 (**wired PR #36**) |
   | IL | IDOT "Getting Around Illinois" RFC viewer | gis1.dot.illinois.gov…/FunctionalClass/MapServer/0 (**wired PR #36**) |
   | OH | ODOT TIMS (`tims.dot.state.oh.us/tims`) | tims…/Functional_Class/MapServer/0 (**wired PR #36**) |

   MN/IL/OH all expose a **bare FHWA 1-7** class field
   (`FUNCTIONAL_CLASS` / `FC` / `FUNCTION_CLASS_CD`) and were wired in
   PR #36 with the Indiana shape (§4.2c-e). MnDOT's same service also has
   a "Federal Adjusted Urban Area" layer (an ACUB analog) if a
   state-native urban-boundary source is ever wanted.

### Column hiding, Sites-only actions, 250 ft buffer, Michigan link (PR #23)

Follow-up after #22 merged (all verified on live services):

4. **Optional columns G-K hidden by default** (Description, Address,
   Category, Costs, Work Completion) on **both** products
   (`modBuild.ApplyProductColumns`), keeping the paste area tight around
   Latitude/Longitude. Reversible; values still flow into KML/CSV/map
   stamps. Column *indices* are unchanged, so verifiers that write by index
   (incl. the FIRMette test writing Category at col 9) still work.
5. **Standard product: all Sites actions moved to the Sites sheet.** Start
   Here's action buttons (Check Roads, Re-run, Photo Links, Export CSV/KML,
   Send-to-AGOL) are now a full button bar on Sites row 1
   (`WriteSitesToolbar` is product-aware: standard = 6 buttons, inspector =
   the 2 shortcuts + its own Start Here toolbar). The row-1 hint text was
   dropped (the yellow `(optional)` headers guide input). Standard Start
   Here is now inputs + Build/Reset only.
6. **Search buffer editable on both products, default 250 ft.**
   `DEFAULT_BUFFER_FEET` and `ACUB_MIN_BUFFER_FEET` both 200 -> **250**; the
   standard Start Here gained the editable buffer cell (`JobBufferFeet` is
   now a shared named range). `verify-classify` re-run at 250 ft — the WI
   STH 52 regression still reads "Non-federal aid - Rural Minor Collector".
7. **Sources sheet opens on "1. SOURCES"** — the title + intro above it were
   removed from `modSources`.
8. **Michigan NFC Map link no longer shows an MPO label.**
   `URL_NFC_MAPVIEW` switched from MDOT's curated NFC/NHS/ACUB webmap (item
   `6a1702b9…`) to the same FEMA-Map-Viewer side-load as IN/WI (MDOT layer
   353 via `url=`, clean `find`/`marker`/`level` pin). The old webmap
   carried an MPO layer whose region labels (e.g. "…Transportation
   Coordinating Initiative") printed over the marker and read as the site's
   name. Trade-off: the curated ACUB overlay isn't shown on that link
   anymore. All three wired states now use the same side-load link shape.

### Per-road distances, ambiguity-aware verdict, Review Reason column, photo-link reorder (PR #24)

Big classify-logic change. The verdict used to be "any detected federal-aid
segment -> red". It now models **the road the point is ON** and flags
*ambiguity* instead of silently guessing.

1. **Per-road distances.** Every road query now fetches geometry
   (`returnGeometry=true&outSR=4326`) and computes a true point-to-polyline
   distance in feet (local equirectangular projection + point-to-segment,
   `modHttp.MinDistanceFt`/`FeatureBlocks`/`PointSegDistM`). The **Road Name**
   column is now a merged, nearest-first `Name (D ft) | Name (D ft)` list
   from all detected roads (state class/route layers + Census TIGER), e.g.
   `S Pitcher St (2 ft) | Sheldon St (197 ft)`.
2. **Primary road drives red/green.** The **closest** class segment decides:
   its class federal-aid -> RED; otherwise green base. This replaces the old
   "any nearby federal road -> red".
3. **Yellow only downgrades green** (never red - "red stays red"). A green row
   becomes yellow ("Review - <reason>") when an ambiguity could actually make
   it federal, with a <=3-word note in the new **Review Reason** column
   (`COL_REVIEWNOTE`, col 20). Reasons, most-specific first: **Second road
   close** (2nd-closest road within 30 ft of the closest AND federal-aid),
   **Nearby FHWA road** (any other detected road is federal-aid), **Urban
   boundary edge** (a Minor Collector whose point is outside the ACUB polygon
   but within the boundary buffer - the only class where urban/rural flips
   the verdict; `DetermineAcub` now distinguishes exact-inside vs
   boundary-buffer vs rural).
4. **Photo links moved right of the results** (user request): imagery links
   (Google Maps..FIRMette Portal) are now cols 21-26, after the results;
   **NFC Map stays on the left** (col 13). Results are cols 14-20
   (contiguous - grey tint + tri-color CF span them). FIRMette/Map Status/AGOL
   unchanged at 27/28/29. The layout is 100% `COL_*`-constant-driven, so
   modBuild/modExport/modMaps needed no edits - only the constants + the
   verifiers' hardcoded indices.

**Build-time compile check (the fix for a recurring trap).** VBA compiles each
module lazily, so a syntax error in a module the build never *executes*
(`modHttp`, `modClassify`) slipped past a "successful" build and only surfaced
(as a modal that hangs headless Excel) at the user's first Check Roads - twice
this round (a mid-module `Private Const`, and a single-line `If ... ElseIf`).
`build\compile-check.ps1` now force-compiles those modules (calls one function
from each, in a background job with a 60 s timeout so a compile-error modal
can't hang) and `build.ps1` runs it after every build. Two VBA rules to keep
in mind (both in §9.3): module-level declarations must precede all procedures,
and a single-line `If` may use `Else` but **not** `ElseIf`.

### Open column -> public app, new AGOL NFC Layer column, hyperlink styling (PR #25)

Split the single per-row map link into two, per user direction:
- **NFC Map / "Open" column** (`COL_NFCMAP`, 13) now opens the state's
  **official public app** (`APP_MI`..`APP_OH` in modConstants - MDOT/INDOT
  Experience apps, MnDOT EMMA, IL Getting Around, OH TIMS; WI's functional-
  class page). App URLs carry NO coordinates (the Experience coordinate
  deep-link mis-navigates, PR #17/#18), so a row opens the app at its default
  extent. PR #23 had wrongly pointed this at a raw Map Viewer side-load - the
  "broke the most important feature" report.
- **New "AGOL NFC Layer" column** (`COL_NFCAGOL`, 30 - appended at the end to
  avoid a 5th full reshuffle) opens the functional-class layer in ArcGIS Map
  Viewer, centered on the point. Michigan uses the curated webmap
  (`URL_NFC_MAPVIEW`, reverted from the 353 side-load) specifically because a
  raw side-load of the time-enabled layer 353 showed a **time slider that hid
  the roads**; the webmap is preconfigured and shouldn't. IN/WI side-load
  their live layer; others get the plain FEMA pin.
- **Link cells now look like links** (blue + underline font on every
  HYPERLINK()-formula column) - Excel doesn't auto-style formula hyperlinks.
- Deleted the "action buttons are on the Sites tab" note (B22) on Start Here.

**Cannot be verified headless:** all of the above are browser behaviors
(Experience-app loading, Map Viewer time slider, side-load rendering). The
build/skeleton confirm the formulas resolve and the Open column points at the
Experience-app URL; the actual in-browser behavior (esp. the Michigan time
slider fix) needs a human click-test.

**"Open Sites on NFC Layer (AGOL)" button** (`modMaps.OpenSitesOnNfcLayer`,
on the Sites toolbar for standard / Start Here for inspector): writes the
verdict-colored KML (red/green/blue, reusing `WriteSitesKml`) and opens the
per-state NFC layer in Map Viewer centered on the first site, then pops
Explorer with the KML for a single drag-drop of all sites onto the layer.
Same untestable-headless caveat.

### Inspector Start Here polish + auto Site # + column tweaks (PR #32)

Inspector-product UX pass, per user direction (standard product unchanged
except the shared bits noted):

1. **No subtitle, no on-sheet disclaimer box.** `BuildStartHereInspector`
   drops the row-3 subtitle and the red `DisclaimerBlock`. The same
   not-authoritative text now shows as a **dialog the first time Check Roads
   runs each session** (`modClassify.ShowDisclaimerOnce`, gated on
   `ProductIsInspector()` + `Not gHeadless` + a module `mDisclaimerShown`
   once-flag). The standard product keeps its on-sheet box. Both surfaces
   read from shared `modBuild.DisclaimerHeaderText()` / `DisclaimerBodyText()`;
   the sentence "It is not an authoritative source for FHWA functional
   classification." was removed from that text.
2. **Two action dropdowns.** The classify + photo actions (Check Roads,
   Re-run Failed Rows, Open Photo Links) collapsed into a **second combo**
   `RR_RoadsPicker` → `RunSelectedRoadsAction`, mirroring the exports picker,
   so the map/FIRMette workflows are the visible hero. `modExportMenu` now
   has a generic `CreatePicker`/`RunPicker` core with two wrappers each.
   `ExportItems()` is **product-aware**: inspector's export dropdown holds only
   the general hand-off exports (CSV, GeoJSON, Send-to-AGOL, Open-on-NFC);
   standard keeps the full list.
3. **FIRMettes and Map Pages are two separate step-by-step hero workflows**
   on inspector Start Here. FIRMettes = Download / Re-run Failed buttons.
   Map Pages = numbered steps 1-6 (input info → Prepare Map Pages → Export
   Sites to KML + screenshot each site with Win+Shift+S saved as
   `site_1.png`… → Insert Map Images → adjust to print area → Export Combined
   Map PDF). KML, FIRMette download/re-run, and combined-map-PDF were
   **promoted from the exports dropdown to dedicated buttons** here.
4. **Search-buffer label** is now "FHWA search buffer (feet)".
5. **Sites columns:** Category (I) and Work Completion (K) now shown;
   Geocode Status (L) hidden (failures still surface in Federal Aid Status).
   Description (G), Address (H), Costs (J) stay hidden.
6. **Site # is auto-numbered by formula** (`SetSiteNumberFormula`): every row
   with a Latitude gets a running 1,2,3…; blank rows stay blank; typing a
   value overrides that row. Tinted grey (computed) and centered.
7. **Sources footer** (both products): a "Created by Caleb Smith / reach out
   to caleb.smith@fema.dhs.gov" mailto line at the very bottom of the sheet.

`verify-skeleton.ps1` updated to match (product-aware button surface, roads
picker present on inspector, disclaimer box asserted on standard only).
Needs the local Windows rebuild + verify pass (§9.2) before hand-off.

---

