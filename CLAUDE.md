# RoadReviewer

A FEMA Public Assistance Site Inspector toolkit. Replaces and consolidates two
prototype workbooks (`GPS Checker - TN updated 3.5.2026.xlsm` and
`Site Inspector Tool 1.xlsm`) into a single, polished, macro-driven Excel
workbook that staff can actually use without training.

Target users: FEMA PA Site Inspectors, most of whom will not use anything
complicated. Target platform: Excel on a hardened government-issued Windows
laptop. No add-ins, no PowerShell, no admin rights — only what is already
installed.

---

## 1. Goal and scope

Three primary activities, each its own workflow with its own entry point but
all reading from one shared Sites table:

1. **Road classification review** — for a point, look up the FHWA functional
   class and whether it falls inside an Adjusted Census Urban Boundary
   (ACUB 2020). Flag anything that is Urban Minor Collector or greater as
   *not eligible* for PA road work.
2. **Pre-disaster imagery review** — open a point in multiple publicly
   available imagery sources (Google, Bing, FEMA Map Viewer, Google Earth,
   Street View, historic imagery where available) so the inspector can
   assess pre-disaster condition.
3. **Firmette & map production** — batch-download FEMA FIRMettes for each
   point and produce per-site location-map pages (with WO/DI/Applicant
   textbox overlays) that can be exported individually or combined.

V2 bonus (deferred): integrate with the existing route-optimization planner
in the other repo.

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

## 3. Requirements for V1

### 3.1 Functional

- **F1. One workbook.** A single `.xlsm` is the deliverable. Proposed
  filename `RoadReviewer.xlsm`. **(decided: single workbook, three workflow
  sheets, one shared Sites table.)**
- **F2. Single source of truth.** All workflows read from one `Sites`
  table on one sheet. The user enters address *or* lat/lon once and never
  retypes it.
- **F3. Three workflows on three sheets**, each with its own clearly
  labelled Start button on a `Home` sheet:
  1. **Classify Roads** — queries the per-state NFC layer + the nationwide
     ACUB layer, writes class / urban-area / eligibility flag back to the
     Sites table.
  2. **Review Imagery** — for the selected row(s), a single button opens
     a curated set of imagery URLs (Google Maps, Google Street View, Bing
     Aerial/Bird's-Eye, FEMA Map Viewer, Esri World Imagery) in default-
     browser tabs.
  3. **Produce Maps & FIRMettes** — batch FIRMette download + map-page
     PDF production (combined or per-site). **MapPages keeps the manual
     screenshot-paste flow in V1**; auto-basemap-capture moves to V2.
- **F4. Address + lat/lon input.** The Sites table accepts either. Empty
  lat/lon with a street address triggers a geocode pass (Census Bureau
  Geocoder — free, no auth, US-only) that fills lat/lon on the row. A
  Geocode Status column shows pass/fail per row. *Lat/lon already in the
  row is never overwritten by the geocoder.*
- **F5. Michigan-only road classification in V1.** Hard-code MDOT NFC
  layer for class + road name. WI / IN / MN / IL / OH are placeholders.
  *(See* [§4 Data sources](#4-data-sources) *.)*
- **F6. ACUB is nationwide.** Use the user-provided AGOL nationwide ACUB
  feature service for all six Region V states (see §4.2). One lookup
  call per row, regardless of state.
- **F7. Eligibility rule.** A row is *PA-ineligible for road work* if any
  intersecting (or nearest-within-150-ft) road segment has a functional
  class of **Urban Minor Collector or greater** inside the relevant
  urban boundary. Implementation:
  - Query the state's NFC layer for class + road name.
  - Cross-check the point against the ACUB polygon layer to determine
    Urban vs Rural.
  - Eligible classes: Rural Local, Rural Minor Collector, Urban Local.
    Everything else = red highlight + "INELIGIBLE" in the Eligibility
    column.
  - **Always run, on every row.** V1 has no visibility into the row's
    PA work Category, so we run eligibility on everything regardless of
    work type. If the tool later integrates with Grants Manager
    (V2+), it could narrow the check to categories where road class
    actually matters — primarily Cat B (Emergency Protective Measures)
    and Cat C (Roads & Bridges), with Cat D (Water Control Facilities),
    Cat F (Public Utilities), and Cat G (Parks/Rec/Other) sometimes
    also touching federal-aid routes.
- **F8. State selector.** Setup sheet has a State dropdown with WI / IN
  / MI / MN / IL / OH. Only MI's NFC layer is wired in V1; selecting
  another state pops a "NFC lookup not yet wired for {state} — ACUB
  still runs" message and the workflow continues with the ACUB-only
  check.
- **F9. NHS column.** Dropped from V1 per user direction. Re-add if
  needed in a later version.
- **F10. Single export across all points.** Provide:
  - One combined Map PDF (already exists — extend it).
  - One KML with all points (already exists).
  - One CSV/XLSX export of the Sites table including all lookup results,
    suitable for handing off to a reviewer.
- **F11. Per-row "Open in map" link.** Every row has a hyperlink that
  drops the user on that exact point in the state's NFC/NHS/ACUB
  Experience app (preserves the prototype's pattern of one click → one
  point).
- **F12. Re-run failed rows only.** A button that re-attempts only rows
  whose status column says "Failed - …".
- **F13. Status panel.** A Status column shows live progress during long
  runs (Application.StatusBar → "Downloading 7 of 42 — Site …"). Already
  present, keep it.
- **F14. WO/DI is per-job.** WO/DI live on the Setup sheet only; the
  inspector can override a single row by manually typing into the row's
  WO/DI cells (always shown, never hidden). No mode toggle.

### 3.2 Non-functional

- **N1. Zero external dependencies** beyond what ships with Office on a
  Windows 10/11 govt laptop. Specifically OK to use:
  `MSXML2.ServerXMLHTTP.6.0`, `ADODB.Stream`, `VBScript.RegExp`,
  `Scripting.FileSystemObject`, `MSForms.DataObject` (via CLSID),
  `Shell` to launch `explorer.exe` / `cmd /c start`.
- **N2. No PowerShell, no add-ins, no installs.** PowerShell is permitted
  *as a last resort* for things the macro cannot do, but the user prefers
  to avoid it.
- **N3. Trust prompts.** The workbook will be opened from
  `OneDrive - FEMA` or SharePoint paths. We need to document how to
  unblock the macro (`Unblock-File` may not be available; right-click →
  Properties → Unblock is the documented path) and we need a one-pager
  for users.
- **N4. Polished.** Consistent column constants, named ranges, defined
  buttons with sensible labels, no orphan macros, no developer artefacts
  (e.g. `Debug.Print`), no commented-out code.
- **N5. Reproducible.** Each workflow can be re-run safely (idempotent
  output: clearing the row's lookup columns before writing new values).
- **N6. Recursive manual verification.** See [§5 Verification plan](#5-verification-plan)
  — each step has an explicit smoke test before we move on.

### 3.3 Out of scope for V1

- WI / IN / MN / IL / OH road-classification queries.
- Address geocoding *(see open question — may move into V1)*.
- Route-optimization integration (V2).
- Embedding ArcGIS-native widgets in Excel.
- Automated screenshot capture from third-party basemaps.
- Authentication to AGOL services.

---

## 4. Data sources

### 4.1 FEMA FIRMette Print service *(unchanged from prototype)*

- Base: `https://msc.fema.gov/arcgis/rest/services/NFHL_Print/MSCPrintB/GPServer/PrintFIRMette`
- Flow: `submitJob?input_lat=…&input_lon=…&Print_Type=FIRMETTE&graphic=PDF&f=pjson`
  → poll `jobs/{jobId}?f=pjson` until `esriJobSucceeded` → fetch
  `jobs/{jobId}/results/OutputFile?f=pjson` → download `url` as PDF.

### 4.2 Michigan road classification (V1)

Source app: MDOT's "NFC, NHS & ACUB" Experience at
<https://experience.arcgis.com/experience/7edd160c205d46b481fcd605bb4c58ce/page/NFC%2C-NHS-%26-ACUB>.
The Experience config itself is not directly fetchable, but the underlying
REST services are publicly readable. **All field names below still need to
be confirmed from the live `?f=pjson` metadata before wiring the VBA —
flagged in the verification plan.**

#### NFC — National Functional Classification (polyline)

- FeatureServer: `https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/353`
- Layer name: `Functional System` (esriLRSLinearEventLayer)
- Spatial reference: WKID 102123 (EPSG 3078 — Michigan GeoRef). The query
  needs `inSR=4326&outSR=4326` when we hand it WGS84 lat/lon.
- Companion layers on the same FeatureServer that may carry the actual
  class code and the route-name text:
  - `…/FeatureServer/364` — "Classification"
  - `…/FeatureServer/543` — "Route System"
- Sample query: `https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/353/query?where=1%3D1&outFields=*&resultRecordCount=1&f=json`
- Metadata: `https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/353?f=pjson`

#### NHS — National Highway System (polyline)

- FeatureServer: `https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/333`
- Layer name: `Nhs`. Key field also called `Nhs` (unique-value renderer).
- Spatial reference: WKID 102123.
- Sample query: `https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/333/query?where=1%3D1&outFields=*&resultRecordCount=1&f=json`

#### ACUB — Adjusted Census Urban Boundary (polygon)

**V1 source: USDOT NTAD "Adjusted Urban Areas" (2020) — nationwide.**
The layer is published as AGOL item `decb7d40c3d540f484dc6925effa9d4b`
("2020 Adjusted Urban Area Boundaries") via UDOT's UPLAN portal, which
re-hosts the USDOT BTS NTAD dataset. **User confirmed nationwide coverage**
visually in the ArcGIS Map Viewer
(`https://www.arcgis.com/apps/mapviewer/index.html?layers=decb7d40c3d540f484dc6925effa9d4b`).
A prior research pass briefly mis-flagged this as Utah-only based on the
host portal name; that inference was wrong.

- **FeatureServer (user-verified):** `https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/NTAD_Adjusted_Urban_Areas/FeatureServer/0`
- Metadata (run from a workstation that can reach `services.arcgis.com`):
  `https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/NTAD_Adjusted_Urban_Areas/FeatureServer/0?f=pjson`
- Sample point-in-polygon query (Detroit, used in verification step §5.1):
  `https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/NTAD_Adjusted_Urban_Areas/FeatureServer/0/query?geometry=-83.045,42.331&geometryType=esriGeometryPoint&inSR=4326&spatialRel=esriSpatialRelIntersects&outFields=*&returnGeometry=false&f=json`
- Hub page (download + schema + sample queries): `https://data-uplan.opendata.arcgis.com/datasets/uplan::2020-adjusted-urban-area-boundaries/about`.
- **Field names — TBD** until the live `?f=pjson` is read from a
  workstation. Expected based on the NTAD schema: an urban-area name
  field (`NAME` or `NAME20`), `UACE20` (urban-area census code),
  `UATYP20` (urbanized area / urban cluster type), and `ALAND20`/`AWATER20`.
  Pinned down in verification step §5.1.

##### Fallback layers (only if the nationwide AGOL item turns out to be unreliable)

- **Michigan MapServer (2010-vintage, "v17a"):** `https://gisagocss.state.mi.us/arcgis/rest/services/OpenData/michigan_geographic_framework/MapServer/5`
  - Fields confirmed: `OBJECTID`, `ACUBCODE` (str4), `ACUBNUM` (int),
    `NAME` (str36, display name), `LABEL` (str60), `TYPE` (str4), `RU`,
    `SQKM`, `SQMILES`, `ACRES`.
  - Sample point-in-polygon query (WGS84 input): `https://gisagocss.state.mi.us/arcgis/rest/services/OpenData/michigan_geographic_framework/MapServer/5/query?geometry=-83.045,42.331&geometryType=esriGeometryPoint&inSR=4326&spatialRel=esriSpatialRelIntersects&outFields=*&returnGeometry=false&f=json`
  - **Caveat:** this is the **2010-vintage** v17a layer. The FHWA-approved
    **2020 ACUB** was reportedly scheduled for MDOT database entry in 2025.
  - Alternate Michigan MapServers carrying the same / similar data:
    - `https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/MapServer/282` (layer `ACUB`)
    - `http://gisp.mcgi.state.mi.us/arcgis/rest/services/MDOT/ACUB/MapServer`
- **Other nationwide candidates** (parked, not used in V1):
  - USDOT BTS NTAD "Urban Areas" — `https://geodata.bts.gov/datasets/usdot::urban-areas/about`
  - FHWA HPMS Adjusted Urban Areas (2020) — `https://www.fhwa.dot.gov/planning/census_issues/urbanized_areas_and_mpo_tma/`
  - US Census TIGER/Line 2020 Urban Areas — `https://www2.census.gov/geo/tiger/TIGER2020/UAC/`

#### Query strategy (Michigan)

Mirroring the TDOT pattern, but with two layers instead of one:

1. **NFC query** — point intersect first. If no hit, fall back to a
   150-foot distance search. Read class + road name.
2. **ACUB query** — point intersect against the polygon layer. If the
   point is inside any ACUB polygon → Urban (record ACUB `NAME` so the
   inspector can see which urban area). Otherwise → Rural.
3. **NHS query** — point intersect (or short fallback). Record Yes/No
   and write to its own column.
4. **Eligibility column** — composed from the NFC class + ACUB result.
   Eligible classes: Rural Local, Rural Minor Collector, Urban Local.
   Anything else → "INELIGIBLE" + red highlight.

The exact field names for the NFC class code and the road/route name are
**TBD** until we read the FeatureServer 353 / 364 / 543 metadata. The
verification plan (§5, step 1) covers this.

### 4.3 Verification map URLs

- Google Maps: `https://www.google.com/maps?q={lat},{lon}`
- Google Street View: `https://www.google.com/maps?q&layer=c&cbll={lat},{lon}`
- Bing Maps: `https://www.bing.com/maps?cp={lat}~{lon}&lvl=18&style=h`
- FEMA Map Viewer: `https://fema.maps.arcgis.com/apps/mapviewer/index.html?find={lon}%2C{lat}&marker={lon},{lat},4326&level=16`
- FEMA FIRMette portal: `https://msc.fema.gov/portal/firmette?latitude={lat}&longitude={lon}`
- MDOT NFC/NHS/ACUB Experience: same URL as above plus a marker (URL parameter format TBD during layer research).

---

## 5. Verification plan

The user wants polished output, so verification is built into the build, not
tacked on afterwards. Every step below ends with a manual smoke test the
developer (or reviewer) actually runs in Excel before moving on.

1. **Layer discovery sanity.** Hit each MDOT REST URL in a browser; confirm
   field names; pick three known coordinates (one clearly Urban Minor
   Collector, one Rural Local inside an ACUB, one Rural Local well outside
   any ACUB). Record expected output.
2. **Skeleton workbook.** Build the Home / Setup / Sites sheets and the
   three workflow sheets. Manual check: open the workbook fresh, confirm
   every button is wired to a sub, no orphan controls.
3. **Sites table.** Wire up the hyperlink formulas and lat/lon validation.
   Manual check: paste in the three known coordinates, confirm every
   hyperlink opens the right map at the right zoom.
4. **Workflow 1 — Classify Roads.** Implement against MDOT layers.
   Manual check: run on the three known coordinates and confirm class,
   road name, urban/rural, eligibility match expectations. Also try a
   point in Tennessee → should report "out of state" not crash.
5. **Workflow 2 — Review Imagery.** Implement the one-click open-many-tabs
   button. Manual check: click on a row, verify every URL opens to the
   correct point.
6. **Workflow 3 — FIRMettes + Maps.** Reuse the prototype's FIRMette and
   MapPages code, but driven from the consolidated Sites table.
   Manual check: batch over 5 sites, confirm 5 PDFs land in the output
   folder with correct names, confirm the combined Map PDF renders
   correctly.
7. **Re-run-failed-rows.** Simulate a failure (kill network briefly), run
   the workflow, then reconnect and click "Re-run failed". Confirm only
   the failed rows are retried.
8. **State selector.** Switch to a non-MI state, confirm the friendly
   "coming in v1.1" message.
9. **Cold-open test.** Close the workbook, copy it to a different
   directory, reopen. Confirm all buttons still work and the macros are
   trusted by Office (this is the most common real-world failure for
   shared `.xlsm` files).
10. **Hand-off review.** Walk a non-developer co-worker through the
    workflow on a real WO. Capture friction points → version 1.0.1
    backlog.

---

## 6. Design ideas for V2+ (not in V1)

These came up reading the prototypes; capturing them so they aren't lost.

- **Auto-capture basemap for MapPages.** Pull a basemap image directly
  from a print service (e.g. ArcGIS World Imagery's `exportImage`
  endpoint, or MDOT's own basemap print service) and drop it into the
  MapPages merged cell automatically, instead of asking the inspector to
  paste a screenshot. Confirmed as a V2 deliverable. Risks to test:
  reachability from a hardened govt laptop, rate limits, attribution
  requirements, and whether the rendered image is high-enough resolution
  for FEMA PA paperwork.
- **Route-optimization integration.** The user has a separate route
  planner. Easiest hand-off: export the Sites table as a CSV / KML / GPX
  that the planner already accepts. Avoid tight coupling for now.
- **AGOL Experience Builder front-end.** For features that are awkward in
  Excel — drawing a polygon search area, side-by-side historical imagery,
  collaborative editing — we can publish the Sites table as a hosted
  feature service and build an Experience app on AGOL. The Excel
  workbook stays the *system of record* (the inspector still owns their
  spreadsheet); the AGOL view is a read-only or write-on-top layer.
  This sidesteps any limitation of "only public maps" because AGOL is
  already licensed for the user's organisation.
- **Historical Google Earth / Nearmap links** for pre-disaster imagery.
- **State expansion.** Wire up the NFC layer for WI / IN / MN / IL / OH.
  Each state's DOT publishes the layer differently; we'll need one
  layer-URL block per state plus a normalisation function that maps each
  state's NFC codes back to the FHWA standard classes. ACUB stays
  nationwide.
- **NHS column.** Surface MDOT's NHS layer (and the equivalent in other
  states) so the inspector can see federal-aid-system status alongside
  functional class.

---

## 7. Repository layout

```
/                                  this README — CLAUDE.md
GPS Checker - TN updated 3.5.2026.xlsm   prototype 1 (reference, do not modify)
Site Inspector Tool 1.xlsm              prototype 2 (reference, do not modify)
RoadReviewer.xlsm                       V1 deliverable (to be built)
docs/                                   user-facing one-pagers (to be added)
```

The two prototype `.xlsm` files are kept in the repo as references. The V1
deliverable is a new workbook so we don't drag along the VBA-stomping
warning from the Site Inspector Tool.

---

## 8. Design decisions (resolved) and remaining open questions

### Resolved (do not relitigate)

1. **ACUB source — nationwide AGOL layer.** The user-provided AGOL item
   `decb7d40c3d540f484dc6925effa9d4b` is the ACUB source for **all six
   Region V states**, not just Michigan. URL and field discovery tracked
   in §4.2.
2. **Address + lat/lon.** V1 accepts both. Census Bureau Geocoder
   (`https://geocoding.geo.census.gov/geocoder/locations/onelineaddress?…`)
   converts address → lat/lon on demand. Lat/lon already in the row is
   never overwritten.
3. **Imagery workflow.** One-click multi-open: a single button on the
   Review Imagery sheet (or context-menu on a Sites row) opens 4–6
   imagery URLs in default-browser tabs for the selected row.
4. **Workbook scope.** Single consolidated `RoadReviewer.xlsm` with a
   Home sheet, one shared Sites table, and three workflow sheets.
5. **Eligibility scope — always run.** V1 has no visibility into the
   row's PA Category, so the road-class check runs on every row. The
   rule is a flag, not a gate: the inspector reads "INELIGIBLE" + class
   on the row and decides whether it applies. If Grants Manager
   integration ever lands (V2+), the check could narrow to Cat B / C /
   D / F / G — the categories that may touch federal-aid routes.
6. **NHS.** Skip in V1. Add later if needed.
7. **MapPages images.** Manual screenshot paste in V1; auto-basemap
   capture is a V2 effort (see §6).
8. **WO/DI.** Per-job only — Setup sheet drives it. Row-level override
   is allowed by editing the row's WO/DI cells directly; no mode toggle.
9. **Output folder default.** Setup sheet pre-fills the output folder
   with `{OneDrivePath}\Desktop\Script\RoadReviewer\{Disaster}\WO{WO}-DI{DI}\`,
   where `{OneDrivePath}` is discovered at workbook open time by
   probing for `%USERPROFILE%\OneDrive - FEMA\`, then
   `%USERPROFILE%\OneDrive\`, falling back to `%USERPROFILE%\Desktop\`
   if neither exists. The folder is created on-demand on first
   FIRMette / Map PDF write. Inspector can override on the Setup sheet.

### Still open

1. **MDOT NFC field names.** §4.2 lists three candidate layers (353
   "Functional System", 364 "Classification", 543 "Route System"). We
   need to read the `?f=pjson` metadata from a machine that can reach
   `mdotgis.state.mi.us` (this build sandbox cannot) to find the exact
   class-code field and road-name field before wiring the VBA. This is
   verification step §5.1; the procedure is documented in
   `docs/probe-mdot-layers.md`.
