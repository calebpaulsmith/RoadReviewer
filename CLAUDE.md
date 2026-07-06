# RoadReviewer

A FEMA Public Assistance road-review toolkit. Replaces and consolidates two
prototype workbooks (`GPS Checker - TN updated 3.5.2026.xlsm` and
`Site Inspector Tool 1.xlsm`) into polished, macro-driven Excel workbooks
that staff can actually use without training.

**Two products are built from one shared `src/` tree** (split decided
2026-07-05; see §7c for the full design):

1. **`RoadReviewer.xlsm`** (standard) — for PDMGs, state/local partners and
   project reviewers. Road classification + photo-source links only; three
   sheets (Start Here / Sites / Sources); inputs are just State, an optional
   Output Folder (blank = next to the workbook) and an optional AGOL webmap
   URL. "Pick the state, paste lat/lon, click Check Roads."
2. **`Site Inspector Review Tool.xlsm`** (inspector) — everything the
   standard product does plus WO/DI/Disaster/Applicant job stamping, batch
   FIRMette download, MapPages and the combined-map PDF. Same three-sheet
   shape.

Target users: FEMA PA staff and partners, most of whom will not use anything
complicated. Target platform: Excel on a hardened government-issued Windows
laptop. No add-ins, no PowerShell, no admin rights — only what is already
installed.

---

## 1. Goal and scope

Three primary activities, each its own workflow with its own entry point but
all reading from one shared Sites table:

1. **Road classification review** — for a point, look up the FHWA functional
   class and whether it falls inside an Adjusted Census Urban Boundary
   (ACUB 2020). Tag anything that is Urban Minor Collector or greater as
   a *federal-aid road*; everything else as *non-federal aid*. The tool
   classifies the road, never the project — the inspector decides what
   federal-aid status means for any given work order.
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

- **F1. One workbook** ~~A single `.xlsm` is the deliverable~~ —
  **superseded 2026-07-05 (§7c): two workbooks are now built from the one
  shared `src/` tree** (`RoadReviewer.xlsm` standard + `Site Inspector
  Review Tool.xlsm` inspector), each with the same three-sheet shape
  (Start Here / Sites / Sources) and one shared Sites table.
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
- **F5. Michigan, Indiana and Wisconsin road classification in V1.**
  Hard-code each state's own NFC layer(s) for class + road name (MDOT for
  MI, INDOT/IndianaMap for IN, WisDOT for WI — WI queries a state-trunk
  layer first and falls back to a local-roads layer). MN / IL / OH are
  still placeholders. *(See* [§4 Data sources](#4-data-sources) *.)*
- **F6. ACUB is nationwide.** Use the user-provided AGOL nationwide ACUB
  feature service for all six Region V states (see §4.2). One lookup
  call per row, regardless of state.
- **F7. Federal-aid rule.** A row is tagged as a *federal-aid road* if
  any intersecting (or nearest-within-150-ft) road segment has a
  functional class of **Urban Minor Collector or greater** inside the
  relevant urban boundary. The tool tags the road, not the project —
  no "eligible"/"ineligible" language anywhere in the output. The
  inspector decides what federal-aid status means for the work order.
  Implementation:
  - Query the state's NFC layer for class + road name.
  - Cross-check the point against the ACUB polygon layer to determine
    Urban vs Rural.
  - Non-federal-aid classes: Rural Local, Rural Minor Collector,
    Urban Local. Federal-aid: Urban Minor Collector or higher.
  - The Federal Aid Status column shows "Federal aid - <class>",
    "Non-federal aid - <class>", or "Review - <reason>". The Sites
    table tints each row light red, green, or yellow accordingly. The
    KML export uses red / green / yellow pushpins for the same three
    buckets; rows that haven't been classified get the default pin.
  - **Always run, on every row.** V1 has no visibility into the row's
    PA work Category, so we run eligibility on everything regardless of
    work type. If the tool later integrates with Grants Manager
    (V2+), it could narrow the check to categories where road class
    actually matters — primarily Cat B (Emergency Protective Measures)
    and Cat C (Roads & Bridges), with Cat D (Water Control Facilities),
    Cat F (Public Utilities), and Cat G (Parks/Rec/Other) sometimes
    also touching federal-aid routes.
- **F8. State selector.** Setup sheet has a State dropdown with WI / IN
  / MI / MN / IL / OH. MI, IN and WI's NFC layers are wired in V1;
  selecting MN, IL or OH pops a "NFC lookup not yet wired for {state} —
  ACUB still runs" message and the workflow continues with the
  ACUB-only check.
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

- MN / IL / OH road-classification queries.
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
REST services are publicly readable. Field names and the
`FunctionalSystem` coded-value domain below were confirmed against the
live `?f=pjson` metadata on 2026-05-22 (verification step §5.1).

**Operational note — MDOT requires a browser User-Agent.** Hitting any
`mdotgis.state.mi.us/arcgis/...` endpoint with the default
`MSXML2.ServerXMLHTTP` UA returns HTTP 403. The VBA HTTP helper must set
a UA header that looks like a browser
(e.g. `Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 …`)
before calling `send`. The NTAD ACUB service (§4.2 ACUB block) has no
such requirement.

#### NFC — National Functional Classification (polyline)

The class code, the road's PR (route identifier) and the route designation
are split across **three companion layers** on the same FeatureServer,
all keyed by `PR` + `PRBmp`/`PREmp` (Michigan's LRS — Physical Reference
+ Begin/End Milepost).

##### Layer 353 — `Functional System` (the class code lives here)

- URL: `https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/353`
- Type: `Feature Layer` (esriGeometryPolyline)
- Spatial reference: WKID 3078 / latestWkid 102123 (Michigan GeoRef).
  Hand it WGS84 with `inSR=4326`.
- displayField: `EventId`
- Confirmed fields (name / type / alias):

  | name | type | alias |
  |---|---|---|
  | `OBJECTID` | esriFieldTypeOID | OBJECTID |
  | `EventId` | esriFieldTypeString | Event ID |
  | `PR` | esriFieldTypeString | PR |
  | `PRBmp` | esriFieldTypeDouble | PR BMP |
  | `PREmp` | esriFieldTypeDouble | PR EMP |
  | `FunctionalSystem` | esriFieldTypeSmallInteger | Functional System |
  | `ProposedFunctionalSystem` | esriFieldTypeSmallInteger | Proposed Functional System |
  | `FieldEstablishDate` | esriFieldTypeDate | Field Establish Date |
  | `RHEstablishDate` | esriFieldTypeDate | RH Establish Date |
  | `RHRetireDate` | esriFieldTypeDate | RH Retire Date |
  | `SystemCreateDate` | esriFieldTypeDate | System Create Date |
  | `SystemModifiedDate` | esriFieldTypeDate | System Modified Date |
  | `UserCreate` | esriFieldTypeString | User Create |
  | `UserModified` | esriFieldTypeString | User Modified |
  | `LocationError` | esriFieldTypeString | Location Error |
  | `Shape__Length` | esriFieldTypeDouble | Shape.STLength() |
  | `GlobalID` | esriFieldTypeGlobalID | GlobalID |
  | `VALIDATIONSTATUS` | esriFieldTypeSmallInteger | Validation status |

- **`FunctionalSystem` coded-value domain** (`LrseFunctionalSystem`):

  | code | name |
  |---|---|
  | 0 | Non-Certified Roadway |
  | 1 | Interstate |
  | 2 | Other Freeway |
  | 3 | Other Principal Arterial |
  | 4 | Minor Arterial |
  | 5 | Major Collector |
  | 6 | Minor Collector |
  | 7 | Local |

  Note: this is the *bare FHWA class code* — it does **not** carry the
  Urban/Rural prefix (unlike the TDOT prototype). Urban vs Rural is
  determined exclusively by the ACUB point-in-polygon check (below).

- **Retired-segment filter is mandatory.** Several intersecting segments
  at any given point can have non-null `RHRetireDate`. Production queries
  must filter `where=RHRetireDate IS NULL` or risk reading historical
  classifications.

##### Layer 364 — `Classification` (feature-type code, e.g. "RD")

- URL: `https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/364`
- Type: `Feature Layer` (polyline), WKID 3078/102123
- displayField: `FeatureClassificationCode`
- Confirmed fields: `OBJECTID`, `EventId`, `PR`, `PRBmp`, `PREmp`,
  `FeatureClassificationCode` (string — at every Detroit/Kalamazoo
  sample point this returned `"RD"`, i.e. "Road"; this layer
  classifies the *kind of feature* — Road vs Trail vs Ramp — not the
  FHWA class), `RHEstablishDate`, `RHRetireDate`, `SystemCreateDate`,
  `SystemModifiedDate`, `UserCreate`, `UserModified`, `LocationError`,
  `Shape__Length`, `GlobalID`. No coded-value domains.
- **Use in V1: skip.** The FHWA class comes from 353 and 364 does not
  carry the urban/rural flag we hoped it might. Keep this URL on file
  for future filtering (e.g. excluding non-road LRS features), but the
  V1 eligibility logic doesn't need it.

##### Layer 543 — `Route System` (route designation = road name source)

- URL: `https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/543`
- Type: `Feature Layer` (polyline), WKID 3078/102123
- displayField: `EventId`
- Confirmed fields: `OBJECTID`, `GlobalID`, `EventId`, `PR`, `PRBmp`,
  `PREmp`, plus three repeats of the route-name tuple:
  - `RouteDesignation`, `RouteNumber`, `RouteBRBLBS`, `RouteBranch`
  - `RouteDesignation2`, `RouteNumber2`, `RouteBRBLBS2`, `RouteBranch2`
  - `RouteDesignation3`, `RouteNumber3`, `RouteBRBLBS3`, `RouteBranch3`

  (plus the standard `RHEstablishDate`/`RHRetireDate`/audit fields).
- **Caveat — only designated trunkline routes are populated.** A
  point-in-polyline (with 150-ft buffer) at downtown Detroit
  (lon=-83.045, lat=42.331) returned **0 features** from layer 543:
  Woodward / Larned / etc. are local streets without an Interstate / US
  / M route designation. The road *name* the inspector cares about
  ("Larned St", "Holmes Rd") is not present in layer 543.
- **Implication for V1.** We get the FHWA class from 353 and the
  *trunkline route name* from 543 when present (e.g. "I-94 BL"). When
  543 returns nothing, the Road Name column will be blank for that row.
  A future enhancement can reverse-geocode the road name from the
  Census Bureau or an OSM source. (Logged for V2 — see §6.)

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
- Type: `Feature Layer` (polygon), spatial reference WKID 4326 (native
  WGS84 — no reprojection needed).
- displayField: `NAME`
- **Confirmed fields** (verification step §5.1, 2026-05-22):

  | name | type | alias |
  |---|---|---|
  | `OBJECTID` | esriFieldTypeOID | OBJECTID |
  | `AREA` | esriFieldTypeDouble | AREA |
  | `UACE` | esriFieldTypeString | UACE |
  | `NAME` | esriFieldTypeString | NAME |
  | `F2020POPUL` | esriFieldTypeDouble | F2020POPUL |
  | `F2020HOUSI` | esriFieldTypeDouble | F2020HOUSI |
  | `F2020_POPD` | esriFieldTypeDouble | F2020_POPD |
  | `state_1` | esriFieldTypeString | state_1 |
  | `Shape__Area` | esriFieldTypeDouble | Shape__Area |
  | `Shape__Length` | esriFieldTypeDouble | Shape__Length |

  No coded-value domains. The urban-area name we surface to the
  inspector lives in **`NAME`** (e.g. `"Detroit, MI"`,
  `"Ann Arbor, MI"`, `"Kalamazoo, MI"`). `UACE` is the 5-digit Census
  urban-area code (used in 2020 release; note no `20` suffix — the
  field names dropped it). `state_1` is the two-letter state
  abbreviation.

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

1. **NFC query** — layer 353. Point intersect first; if no hit, fall
   back to a 150-foot distance search. Always filter
   `where=RHRetireDate IS NULL` to exclude retired segments. Read
   `FunctionalSystem` (smallint → look up label via the coded-value
   domain above) and `PR`. The road name itself is **not** carried on
   layer 353 — see step 2.
2. **Road-name query** — layer 543 with the same point-intersect /
   150-ft fallback. Concatenate any non-null
   `RouteDesignation{,2,3}` + `RouteNumber{,2,3}` tuples into a
   display string. Expect this to be **blank for local streets**
   (only trunkline routes are populated); that's fine.
3. **ACUB query** — point intersect against the NTAD polygon layer.
   If the point is inside any ACUB polygon → Urban (record `NAME`,
   `UACE`, `state_1`). Otherwise → Rural.
4. **NHS query** — V1 skips this (per F9). Layer 333 URL kept on file
   for V1.1.
5. **Federal Aid Status column** — composed from the NFC class + ACUB
   result. Non-federal aid: (Urban + Local), (Rural + Local),
   (Rural + Minor Collector). Federal aid: everything else. The
   smallint -> tag table in VBA terms:

   | `FunctionalSystem` | ACUB hit? | Result |
   |---|---|---|
   | 7 (Local) | yes (Urban) | Non-federal aid - Urban Local |
   | 7 (Local) | no (Rural) | Non-federal aid - Rural Local |
   | 6 (Minor Collector) | no (Rural) | Non-federal aid - Rural Minor Collector |
   | 6 (Minor Collector) | yes (Urban) | Federal aid - Urban Minor Collector |
   | 5 / 4 / 3 / 2 / 1 | either | Federal aid - <Urban/Rural class> |
   | 0 (Non-Certified) | either | Review - non-certified class, check manually |

   When multiple segments are returned (intersection within 150 ft),
   the row is tagged federal-aid if *any* returned segment is.

#### Confirmed test coordinates (verification §5.1)

The following three known points were verified live against the MDOT
353 + NTAD ACUB services on 2026-05-22. Use them as the smoke-test set
for workflow 1.

| # | Expected outcome | lat | lon | NFC class | ACUB |
|---|---|---|---|---|---|
| 1 | **Federal aid** — Urban Minor Collector | `42.28536` | `-85.57025` | `6` (Minor Collector), single segment, PR=`0006904` | `Kalamazoo, MI` (UACE=`43723`) |
| 2 | Non-federal aid — Urban Local | `42.6911` | `-84.5360` | `7` (Local), single segment, PR=`0343402` (Holmes Rd corridor) | `Lansing, MI` (UACE=`47719`) |
| 3 | Non-federal aid — Rural Local | `44.2700` | `-83.5200` | `7` (Local), single segment, PR=`1257508` (Iosco County, near Tawas City) | none — point is rural |

Note on §5 step 1's third bullet ("Rural Local inside an ACUB"): a
class-7 polyline that intersects an ACUB polygon is by definition
*Urban* Local under the eligibility rule, not Rural Local. The intent
is "a class-7 (Local) road inside an ACUB" — that's test case #2
above.

### 4.2a Indiana road classification (V1)

Source: INDOT's public GIS platform at `gisdata.in.gov` (the "IndianaMap"
hosted-services back end), confirmed live 2026-07-01.

#### Layer `LRSE_Functional_Class` — the class code

- URL: `https://gisdata.in.gov/server/rest/services/Hosted/LRSE_Functional_Class/FeatureServer/22`
- Type: `Feature Layer` (esriGeometryPolyline), spatial reference WKID
  26916 (NAD83 UTM 16N). Hand it WGS84 with `inSR=4326`, same pattern as
  MDOT.
- displayField: `from_date`
- Confirmed fields: `from_date`, `to_date` (Date), `event_id` (String),
  `route_id` (String — LRS event key), `from_measure`/`to_measure`
  (Double), `record_status` (SmallInteger), **`functional_class`**
  (SmallInteger), `created_by`, `date_created`, `edited_by`,
  `date_edited`, `locerror` (String), `objectid` (OID),
  `date_attr_effective` (Date), `globalid` (GlobalID),
  `SHAPE__Length`.
- **`functional_class` coded-value domain (`dFunctionalClass`)** — same
  structure and numbering as MDOT's `LrseFunctionalSystem` domain, no
  urban/rural embedded:

  | code | name |
  |---|---|
  | 1 | Interstate |
  | 2 | Principal Arterial - Other Freeways/Expressways |
  | 3 | Principal Arterial - Other |
  | 4 | Minor Arterial |
  | 5 | Major Collector |
  | 6 | Minor Collector |
  | 7 | Local |

- **`record_status` coded-value domain (`dRecordCode`)** — Indiana's
  analog to Michigan's `RHRetireDate IS NULL` filter:

  | code | name |
  |---|---|
  | 0 | Work In Progress |
  | 1 | Proposed |
  | 2 | Withdrawn |
  | 3 | Rejected |
  | 4 | Accepted |
  | 5 | Active |
  | 6 | Replaced |
  | 7 | Retired |

  Production queries filter **`where=record_status=5`** (Active). All
  live test points below returned status 5.

#### Layer `Road_Centerlines_of_Indiana_2021` — road name (separate, un-keyed)

- URL: `https://gisdata.in.gov/server/rest/services/Hosted/Road_Centerlines_of_Indiana_2021/FeatureServer/15`
- Type: `Feature Layer` (esriGeometryPolyline), WKID 26916, displayField
  `st_name`.
- Confirmed fields: `st_name`, **`st_full`** (best display value, e.g.
  `"MERIDIAN ST"` — this is what RoadReviewer reads), directional/type
  modifier components, and a `roadclass` string field (returned null in
  every sample — not reliable, not used).
- **This layer is not LRS-keyed to `LRSE_Functional_Class`** — there is
  no shared join field, unlike MDOT where 353/543 both key off
  `PR`/`PRBmp`/`PREmp`. Road name is resolved via a fully separate
  point-intersect query (same exact-point-then-150-ft-buffer pattern).
  Indiana's functional-class layer carries no name field at all — not
  even a trunkline-only one like MDOT 543 — so this comes back blank
  more often than Michigan's Road Name column; Census TIGER (already
  wired for every state) backs it up.

#### UA / header quirks

None observed. `gisdata.in.gov` answered plain `f=json` GET requests
without a browser User-Agent (unlike `mdotgis.state.mi.us`). RoadReviewer
sends the browser UA on every request regardless (`modHttp.HttpGetText`),
which is harmless here.

#### Confirmed test coordinates (verified live 2026-07-01)

| # | Expected outcome | lat | lon | `functional_class` | `record_status` |
|---|---|---|---|---|---|
| 1 | Higher-class road | `39.7684` | `-86.1581` | `6` (Minor Collector), downtown Indianapolis / Monument Circle area, route_id `549095041900000R1` | `5` (Active) |
| 2 | Higher-class road | `39.4234` | `-86.7628` | `3` (Principal Arterial - Other), near Martinsville, route_id `20000002310000001` | `5` (Active) |
| 3 | Local road | `39.9876` | `-86.0128` | `7` (Local), route_id `52901903520000001` — originally logged as "rural Hancock County", but the 2026-07-05 verify-classify run showed the point is at the Marion/Hamilton county line INSIDE the `Indianapolis, IN` ACUB polygon, so the correct verdict is "Non-federal aid - Urban Local" (the class code was the live-verified part) | `5` (Active) |

### 4.2b Wisconsin road classification (V1)

Source: two companion AGOL-hosted feature services owned by WisDOT
(`services5.arcgis.com`, account `wisnipsvki_WisDOT`), both explicitly
built for damage-assessment use — confirmed live 2026-07-01. Same
hosting pattern as the nationwide NTAD ACUB layer (§4.2), not a
state-hosted server like MDOT's — so no browser-UA workaround needed.

Wisconsin is the only wired state that needs **two layers queried in
sequence**: the State Trunk Network covers interstates/state highways
with a clean bare FHWA code; the Local Road Network snapshot covers
everything else (county/municipal/town roads) but encodes urban/rural
into its own class code instead of carrying a bare FHWA number.

#### Layer `FFCL_gdb/3` — State Trunk Network (primary; state highways/interstates)

- URL: `https://services5.arcgis.com/0pgGLzT0Nh7FVjon/arcgis/rest/services/FFCL_gdb/FeatureServer/3`
  ("Extract of the Functional Class - State Trunk Network for IDA
  Application")
- Type: `Feature Layer` (esriGeometryPolyline), spatial reference WKID
  102100/3857 (Web Mercator). Hand it WGS84 with `inSR=4326`; the
  service reprojects automatically.
- Confirmed fields: `OBJECTID`, `FED_FUNC_CLS_ID`, `RWLK_ID`, `FC_CD`
  (WisDOT's own internal code), `FC_ABBR_DESC`, `FC_DESC`,
  `LAST_CHANGED_BY`, `LAST_CHANGED_ON`, **`FED_FC_CD`** (String — bare
  FHWA class code, `"1"`–`"7"`), `URB_TYPE` (String: `"Urban"` /
  `"Rural"` / null — **not used**; RoadReviewer keeps ACUB as the single
  urban/rural source of truth for every state, per §4.2), `FED_FC_DESC`,
  `FC_TYPE`, **`HWYTYPE`** (STH/USH/IH/OFF), **`HWYNUM`**, **`HWYDIR`**,
  `SEGLEN`, `DIV_STATUS`, `FROM_OFFSET`, `TO_OFFSET`, `Shape__Length`.
- **`FED_FC_CD` values (confirmed live via `returnDistinctValues=true`)**
  — standard FHWA 1-7, identical numbering to Michigan and Indiana:

  | FED_FC_CD | FED_FC_DESC |
  |---|---|
  | 1 | Principal Arterial - Interstate |
  | 2 | Principal Arterial - Freeways and Expressways |
  | 3 | Principal Arterial - Other |
  | 4 | Minor Arterial |
  | 5 | Major Collector |
  | 6 | Minor Collector |
  | 7 | Local |
  | *(null)* | unclassified / off-network (ramps, etc.) |

- Road name is built from `HWYTYPE` + `HWYNUM` + `HWYDIR` (e.g. `"STH 32
  S"`, `"USH 18 E"`).
- No retired-segment filter needed — this is described as a point-in-time
  snapshot/extract service (no version/retire-date field in the schema),
  unlike MDOT's editable production layer.

#### Layer `WI_Local_Roads_Flood_Damage_Assessment_Snapshot/1` — Local Road Network (fallback; everything else)

- URL: `https://services5.arcgis.com/0pgGLzT0Nh7FVjon/arcgis/rest/services/WI_Local_Roads_Flood_Damage_Assessment_Snapshot/FeatureServer/1`
- Type: `Feature Layer` (esriGeometryPolyline). Same WISLR-derived ~90-field
  schema as WisDOT's general-purpose `WISLR_Functional_Road_Classification`
  hosted layer, filtered to the local-road network — state highways
  return null functional-class fields here (confirmed at WIS 181 and
  I-41), which is why the State Trunk layer above is tried first.
- Relevant fields: **`ST_LABL_NM`** (road name — populated for local
  streets, better coverage here than MDOT's trunkline-only layer 543),
  `FNCT_CLS_TYCD` / **`FNCT_CLS_CTGY_TYCD`** (WisDOT's own class scheme —
  see domain below), `FEDUA_TYCD` (3-digit urban-area code, `'000'` =
  rural/no urban area, e.g. `'057'` = Milwaukee urban area — not used,
  same ACUB-is-source-of-truth reasoning as `URB_TYPE` above),
  `NHS_CLS_TYCD` (`'NHS'`/`'NON'` — unused, NHS is out of scope per F9).
- **`FNCT_CLS_CTGY_TYCD` domain** — confirmed live via the layer's
  renderer `uniqueValueInfos` (no formal coded-value domain is published
  in the service metadata, so this was read off the map-symbology
  config instead):

  | value | label |
  |---|---|
  | 10 | Rural Principal Arterial |
  | 60 | Urban Principal Arterial |
  | 20 | Rural Minor Arterial |
  | 86 | Urban Minor Arterial Other |
  | 30 | Rural Major Collector |
  | 96 | Urban Collector Other |
  | 40 | Rural Minor Collector |
  | 45 | Rural Local |
  | 97 | Urban Local |

  Unlike the state-trunk layer's `FED_FC_CD`, **this field embeds
  urban/rural directly into the code** (like Tennessee's `FUNC_CLASS` in
  the prototype) instead of carrying a bare FHWA number.
  `WisconsinLocalCategoryToFhwa()` in `modConstants.bas` strips the
  urban/rural digit back out to a bare FHWA 1-7 code before this feeds
  into the shared `FederalAidVerdict()` logic. One quirk: `96` ("Urban
  Collector Other") doesn't distinguish Major from Minor Collector.
  That split only changes the federal-aid verdict for *rural* collectors
  (rural major = federal aid, rural minor = not) — every *urban*
  collector is federal aid regardless of major/minor (§4.2's eligibility
  table), so mapping `96` to Major Collector (5) is safe either way and
  never produces a wrong verdict.
- No retired-segment filter needed (same "point-in-time snapshot" pattern
  as the state-trunk layer).

#### Query strategy (Wisconsin)

1. Point-intersect (then 150-ft/configured-buffer fallback) against
   `FFCL_gdb/3`. If any segment is returned, read `FED_FC_CD` directly
   as the bare FHWA code and stop — state highways are always resolved
   here.
2. Only if step 1 returns nothing: point-intersect (then buffer
   fallback) against `WI_Local_Roads_Flood_Damage_Assessment_Snapshot/1`,
   decode `FNCT_CLS_CTGY_TYCD` via `WisconsinLocalCategoryToFhwa()`.
3. ACUB (urban/rural) runs exactly as it does for every other state —
   independently of whichever WI layer answered the class query.

#### UA / header quirks

None. Both layers are AGOL-hosted `services5.arcgis.com` feature
services — default UA (or any UA) returns 200, same as the nationwide
ACUB layer. RoadReviewer sends its browser UA on every request
regardless; harmless here.

#### Confirmed test coordinates (verified live 2026-07-01)

| # | Expected outcome | lat | lon | Layer / result |
|---|---|---|---|---|
| 1 | Higher-class, urban | `43.0389` | `-87.9065` | State Trunk: `FED_FC_CD='4'` (Minor Arterial), `URB_TYPE='Urban'` — E Wisconsin Ave / USH 18 / STH 32, downtown Milwaukee |
| 2 | Higher-class, rural | `45.4711` | `-89.7345` | State Trunk: `HWYTYPE='STH'`, `HWYNUM='86'`, `FED_FC_CD='5'` (Major Collector), `URB_TYPE='Rural'` — near Tomahawk, WI |
| 3 | Local road, rural | ~`46.65` | ~`-90.86` | Local Roads (state-trunk layer returns nothing here): `FNCT_CLS_CTGY_TYCD=45` (Rural Local), `FEDUA_TYCD='000'` — Ballard Rd, Washburn County |

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
4. **Workflow 1 — Classify Roads.** Implement against MDOT (MI), INDOT
   (IN) and WisDOT (WI) layers. Manual check: run on each state's known
   coordinates (§4.2, §4.2a, §4.2b) and confirm class, road name,
   urban/rural, federal-aid status match expectations. Also try a point
   in Tennessee → should report "out of state" not crash.
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
8. **State selector.** Switch to an unwired state (MN/IL/OH), confirm
   the friendly "not yet wired" message. MI/IN/WI must NOT show it.
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
- **State expansion.** MI, IN and WI are wired (§4.2, §4.2a, §4.2b).
  Wire up the NFC layer for MN / IL / OH next. Each state's DOT publishes
  the layer differently; we'll need one layer-URL block per state plus a
  normalisation function that maps each state's NFC codes back to the
  FHWA standard classes (see `WisconsinLocalCategoryToFhwa` in
  `modConstants.bas` for a worked example — WI's local-roads layer embeds
  urban/rural into its own class code and needs unpacking). ACUB stays
  nationwide.
- **NHS column.** Surface MDOT's NHS layer (and the equivalent in other
  states) so the inspector can see federal-aid-system status alongside
  functional class.

---

## 7. Repository layout

```
/                                  this README — CLAUDE.md
.github/workflows/pages.yml       deploys web/ to GitHub Pages on push to main (§7b)
GPS Checker - TN updated 3.5.2026.xlsm   prototype 1 (reference, do not modify)
Site Inspector Tool 1.xlsm              prototype 2 (reference, do not modify)
RoadReviewer.xlsm                       Standard product (PDMGs/partners/reviewers) — committed;
                                          rebuild via `build\build.ps1` after any src/ change
Site Inspector Review Tool.xlsm         Inspector product (full toolkit) — committed; built from
                                          the SAME src/ tree by the same script (§7c)
src/                                    Shared VBA source (importable .bas modules, both products)
  modConstants.bas                      product ids, sheet names, column indices, URLs,
                                          FunctionalSystem domain, FIRMette poll constants,
                                          MapPages layout constants
  modUtil.bas                           shared helpers + gHeadless/gTrace plumbing + product
                                          identity (SetProduct / ProductIsInspector / ProductTitle)
  modHttp.bas                           browser-UA GET + narrow JSON extraction (VBScript.RegExp)
                                          + HttpDownloadPdf (binary PDF via ADODB.Stream)
  modBuild.bas                          BuildWorkbook — product-branched Start Here + Sites
                                          (toolbar row, tints, product column-hiding)
  modSources.bas                        Sources sheet — per-state citations + quirks (§7c)
  modClassify.bas                       Check Roads — geocode + NFC + ACUB + route name +
                                          federal-aid verdict, re-run-failed (F7/F12)
  modGeocode.bas                        GeocodeRow helper — Census one-line geocoder, called
                                          per-row from Check Roads (F4; no standalone button)
  modImagery.bas                        open curated photo-link set (incl. Google Earth) for
                                          selected rows
  modMaps.bas                           Inspector workflow 3 — FIRMette download (FEMA GP
                                          service), MapPages layout, ExportCombinedMapPdf,
                                          KML export, output-folder resolution
  modExport.bas                         Sites table → CSV with resolved link URLs, product-
                                          filtered columns (F10)
build/                                  Local assembly + verification scripts (not for end users)
  build.ps1                             COM-driven build; `-Product Standard|Inspector|Both`
                                          (default Both) → RoadReviewer.xlsm + Site Inspector
                                          Review Tool.xlsm at the repo root (stages SaveAs in
                                          %TEMP% because the repo dirs carry an Everyone-Deny
                                          ACE that breaks Excel's SaveAs temp-file pattern)
  BuildHelper.bas                       Build-time-only module; sets gHeadless + traps errors
                                          to %TEMP%\RoadReviewer_build_error.txt, removed
                                          before save
  verify-skeleton.ps1                   §5.2 + §5.3 — sheets, buttons, named ranges, headers,
                                          hyperlink formulas, decimal validation; product-aware
                                          (auto-detects RR_Product, asserts each product's
                                          button/name/hidden-column surface) — run once per built file
  verify-classify.ps1                   §5.4 — test coords for MI/IN/WI against live MDOT/INDOT/WisDOT + NTAD,
                                          plus an address-only row exercising Check Roads' auto-geocode
  verify-rerun-and-state.ps1            §5.7 + §5.8 — re-run-failed-rows + state selector
  verify-firmette-maps.ps1              Workflow 3 — DownloadFirmettes + PrepareMapPages
                                          + ExportCombinedMapPdf against live FEMA GP
  dump-prototype.ps1                    Extracts the prototype VBA modules to build/prototype-vba/
                                          for reference (not version-controlled)
  verify-web-core.mjs                   web prototype — executes web/index.html's rr-core
                                          <script> block headless (Node + curl) against the live
                                          services, asserting the §4.2/§4.2a/§4.2b test coords
  web-tests/                            Playwright-driven verifiers for the web prototype's
                                          browser-only features (real browser + canvas, so they're
                                          separate from verify-web-core.mjs's curl-only approach)
    verify-pdf-report.mjs               classifies a point, clicks "Download PDF Report", checks
                                          the resulting PDF's structure (page count, embedded images)
    verify-review-ui.mjs                map name labels, click-to-zoom, Prev/Next stepping, on-map
                                          source layers + legend, sources.html, FIRMette ZIP batch
                                          (ZIP validated with Python zipfile incl. CRCs)
    fixtures/                           real captured MDOT + ACUB responses used to stub the
                                          network (see each script's header comment for why)
web/                                    FHWA Road Checker (provisional) — static web prototype (§7b)
  index.html                            single-file page: paste coords → auto-classify → map pins
                                          → site-by-site review (zoom, Prev/Next, on-map source
                                          layers + legend) → optional "Download PDF Report" and
                                          "Download FIRMettes (ZIP)"
  sources.html                          per-state data-source citations: org, service URL, exact
                                          layer names, fields read, schema quirks
  vendor/leaflet/                       Leaflet 1.9.4 vendored locally (no CDN calls)
  vendor/jspdf/                         jsPDF 2.5.2 UMD build vendored locally (no CDN calls)
  README.md                             privacy model, hosting, PDF report design, verification
docs/
  probe-mdot-layers.md                  how to re-run the §5.1 schema probe locally
  probe.py                              stdlib probe script for the four FeatureServers
  build-and-import.md                   how to assemble RoadReviewer.xlsm from src/ on a Windows laptop
```

The two prototype `.xlsm` files are kept in the repo as references. The V1
deliverable is a new workbook so we don't drag along the VBA-stomping
warning from the Site Inspector Tool. The workbook **is committed at the
repo root** so it's accessible to inspectors who clone the repo, and is
rebuilt on the user's Windows laptop from `src/` (see
`docs/build-and-import.md`) — the cloud build environment has no Excel,
so changes to `src/` need a local rebuild + commit before the .xlsm
reflects them.

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
  federal-aid buckets as the Sites table and KML export. A visible
  network log lists every request the page makes; Leaflet is vendored
  locally so there are no CDN calls. Exports: CSV download +
  copy-for-Excel TSV, both generated in the browser.
- **Same logic as Excel.** `web/index.html`'s `<script id="rr-core">`
  block is a hand-port of `modClassify.bas`/`modConstants.bas` (same
  where-clauses, fallback buffers incl. the 200-ft ACUB floor, verdict
  table, state gate). The two must be kept in sync by hand — there is no
  shared source between VBA and JS. `build/verify-web-core.mjs` executes
  that exact script block headless against the live services and passed
  on all §4.2/§4.2a/§4.2b test coordinates (2026-07-02), including the
  two increment-4 WI regression points; a Playwright DOM smoke (paste →
  table rows + colored markers, stubbed services) also passed. MDOT
  throws occasional transient 503s; failed rows aren't cached and get a
  per-row retry link (web analog of F12).
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

### Front-page disclaimer, boundary handling, citation page, version stamp (PR #19)

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
   (`"PR #19"`), shown as a small grey label at the bottom of Start Here
   (`modBuild.VersionLabel`) and in the Sources footer, so a shared copy
   is traceable to the build/PR it came from. Bump it each release.
   `verify-skeleton.ps1` asserts the disclaimer text, the eligibility
   clause, the `PR #` label, and the `BOUNDARY ROADS` caveat.

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

10. **MDOT NFC field names — confirmed 2026-05-22 (verification §5.1).**
    Read live from `mdotgis.state.mi.us` via the procedure in
    `docs/probe-mdot-layers.md`. Class comes from layer **353**, field
    `FunctionalSystem` (smallint, coded-value domain 0–7), filtered
    `where=RHRetireDate IS NULL`. Layer 364 is feature-type only ("RD")
    and skipped. Trunkline route name comes from layer **543**
    (`RouteDesignation`/`RouteNumber`), blank for local streets. ACUB is
    the NTAD layer's `NAME` field (native WGS84). Full schema, eligibility
    table, and three live-verified test coordinates are in §4.2.
11. **INDOT + WisDOT NFC field names — confirmed live 2026-07-01.**
    `gisdata.in.gov` and `services5.arcgis.com` are both reachable from
    this repo's cloud sandbox (unlike `mdotgis.state.mi.us`), so this
    round was verified directly rather than via the local-workstation
    probe workflow. Indiana's `functional_class` (LRSE_Functional_Class,
    layer 22) and Wisconsin's `FED_FC_CD` (state-trunk layer) both share
    Michigan's bare-FHWA-1-7-code shape with no urban/rural embedded.
    Wisconsin's local-roads fallback layer is the one exception — its
    `FNCT_CLS_CTGY_TYCD` embeds urban/rural into the code itself, handled
    by `WisconsinLocalCategoryToFhwa()` in `modConstants.bas`. Full
    schemas, eligibility mapping, and live-verified test coordinates for
    both states are in §4.2a and §4.2b.

### Still open

One cosmetic item unchanged from before: the MDOT NFC/NHS/ACUB
Experience-app marker URL parameter format (§4.3) is still TBD — only
affects the per-row "Open in map" deep link (F11), not the classification
logic. Can be pinned down during the §5.4 smoke test.

The "NFC Map" Sites column (col 19 since the §7c reorder, formerly labeled
"MDOT NFC Map") is state-aware (§9.3b): it dispatches on Setup's State dropdown between
`URL_NFC_MAPVIEW` (MI's curated FEMA webmap), `URL_NFC_MAPVIEW_IN`, and
`URL_NFC_MAPVIEW_WI` (the latter two side-load the state's own live NFC
FeatureServer into the plain FEMA Map Viewer rather than a curated
webmap, since no combined IN/WI NFC+ACUB webmap exists to point at).
Unverified in a browser - built and reasoned through from the same
Map Viewer URL parameters already proven for `URL_FEMAVIEW`, but the
`url=`-sideload behavior itself hasn't been click-tested. Worth
confirming during the §5.5/§5.4 smoke pass that the IN/WI links actually
render the layer (they'll still open to the correct pin+zoom even if the
layer overlay doesn't render, so this degrades gracefully either way).

One thing blocking *trusting* Indiana/Wisconsin output on a real WO (not
blocking the code itself): §4.2a/§4.2b's schemas were confirmed by
fetching each service's live JSON directly, not by running RoadReviewer's
actual VBA HTTP stack against them from Excel. Run the local
`verify-classify.ps1` (§9.2, now covers all three wired states) at least
once before relying on IN/WI classifications for real inspection work.

---

## 9. Developer notes — rebuild workflow and VBA gotchas

Reference material for whoever is editing `src/` next. Lookup-style; the
narrative version of how these bugs got found is in §7a.

### 9.1 Rebuild the workbooks after a `src/` change

`build\build.ps1` builds **both products by default**
(`-Product Standard|Inspector|Both`, default `Both`) into the repo root:
`RoadReviewer.xlsm` + `Site Inspector Review Tool.xlsm`. Close any open
Excel first — the files are locked while open.

```powershell
# from anywhere - builds both products
& "C:\Users\caleb\OneDrive\Desktop\Scripts\RoadReviewer\build\build.ps1"

# then commit the refreshed .xlsm files
cd "C:\Users\caleb\OneDrive\Desktop\Scripts\RoadReviewer"
git add RoadReviewer.xlsm "Site Inspector Review Tool.xlsm" <other src files>
git commit -m "Rebuild workbooks: <what changed in src/>"
git push origin main
```

The build, per product:

1. Starts a hidden Excel via COM (one instance for the whole run).
2. Imports every `.bas` from `src/` plus `build\BuildHelper.bas`.
3. Calls `SetProduct` (bakes the hidden `RR_Product` name, §7c) then
   `BuildWorkbookSafe`, which flips `gHeadless = True` and invokes
   `BuildWorkbook`. The Safe wrapper writes any error to
   `%TEMP%\RoadReviewer_build_error.txt` and re-raises so the COM host
   sees a real exception instead of Excel sitting in VBE break mode.
4. Removes `BuildHelper` from the project before save (build-time-only).
5. SaveAs into %TEMP% staging as `xlOpenXMLWorkbookMacroEnabled` (52),
   then `Copy-Item -Force` into the repo root (ACL workaround, §9.4).
6. Quits Excel cleanly after the last product.

Takes ~20s for both on a typical laptop. Use `-Product` for one product
and `-OutDir` for a different destination folder.

### 9.2 Verifier suite

Every script in `build/` flips `gHeadless` on, runs a target workflow
against a built workbook (`-XlsmPath`), and asserts results. The classify
scripts work against either product (the classify path is shared); the
FIRMette/MapPages scripts need the **inspector** build (the standard
product has no FIRMette buttons or WO/DI named ranges).

| Script | What it covers | Run against | Network? |
|---|---|---|---|
| `verify-skeleton.ps1` | §5.2 + §5.3 — sheets, buttons (every OnAction resolves), product button/named-range surface, Sites headers (row 2), hidden inspector-only columns, toolbar buttons, hyperlink formulas (incl. Google Earth), lat/lon decimal validation, verdict conditional formatting, Sources content | each product | no |
| `verify-classify.ps1` | §5.4 — Check Roads against the §4.2/§4.2a/§4.2b test coords for MI, IN and WI + an address-only auto-geocode row | standard (or either) | MDOT + INDOT + WisDOT + NTAD + Census |
| `verify-rerun-and-state.ps1` | §5.7 + §5.8 — state=MN gates NFC (MI/IN/WI must NOT gate); ReRunFailedRows only retries `Failed - ` rows | standard (or either) | MDOT + NTAD |
| `verify-firmette-maps.ps1` | Inspector workflow 3 — DownloadFirmettes + PrepareMapPages + ExportCombinedMapPdf on one Kalamazoo site | inspector | FEMA Print FIRMette GP |
| `verify-blank-wodi.ps1` | PR #5 — empty WO/DI produces clean filenames + stamps (no dangling `WO `, ` DI`, or `WO #` line) | inspector | FEMA GP |

Run the whole suite from a clean state:

```powershell
$repo = "C:\Users\caleb\OneDrive\Desktop\Scripts\RoadReviewer"
& "$repo\build\build.ps1"                  # builds both products
$std = "$repo\RoadReviewer.xlsm"
$ins = "$repo\Site Inspector Review Tool.xlsm"
& "$repo\build\verify-skeleton.ps1"        -XlsmPath $std
& "$repo\build\verify-skeleton.ps1"        -XlsmPath $ins
& "$repo\build\verify-classify.ps1"        -XlsmPath $std
& "$repo\build\verify-rerun-and-state.ps1" -XlsmPath $std
& "$repo\build\verify-firmette-maps.ps1"   -XlsmPath $ins
& "$repo\build\verify-blank-wodi.ps1"      -XlsmPath $ins
```

Trace output from any verifier lands in
`%TEMP%\RoadReviewer_classify_trace.txt` (or `_w3_trace.txt` for
Workflow 3) when `SetTrace` is on — use it to see exactly which HTTP
call was in flight if a workflow hangs.

### 9.3 VBA gotchas that bit this codebase

A lookup table for the bugs we discovered the first time `BuildWorkbook`
ran in Excel. Each one is also called out at its fix site in `src/`.

**`Worksheet.DisplayGridlines` doesn't exist.** Gridlines are a `Window`
property, not a `Worksheet` property. Use `HideGridlines(ws)` in modBuild,
which activates the sheet and flips `ActiveWindow.DisplayGridlines` with
`On Error Resume Next` for the headless-COM case where `ActiveWindow`
is `Nothing`.

**Module-level `Public` declarations must precede every Sub.** VBA throws
*"Only comments may appear after End Sub"* if you put `Public foo As
String` between two procedure bodies. All Public state lives at the top
of `modUtil`.

**`global` is a VBA reserved word** (synonym for `Public`). Using it as a
parameter name (`Private Function NewRegex(..., ByVal global As Boolean)`)
imports clean, then JIT compilation at first call throws *"Sub or
Function not defined"* on every caller. The function is renamed to
`isGlobal`.

**`line` is also reserved.** Used by `Line Input #` and the `Line`
method on Shape. Same JIT-failure pattern. `TraceLine` takes
`ByVal txt As String`, not `ByVal line As String`.

**ArcGIS JSON has case-sensitive attribute keys.** Default
`NewRegex(..., ignoreCase:=False)`. A case-insensitive
`FirstString("NAME")` matches the lower-case `"name":"OBJECTID"` field-
metadata entry before the actual `"NAME":"Kalamazoo, MI"` attribute,
breaking the ACUB-name column.

**MDOT requires a browser User-Agent.** `mdotgis.state.mi.us` returns
HTTP 403 to the default `MSXML2.ServerXMLHTTP` UA. All HTTP traffic
goes through `HttpGetText` / `HttpDownloadPdf` in `modHttp`, which set
a Chrome-style UA via `BROWSER_UA`. NTAD ACUB (`services.arcgis.com`)
doesn't need it.

**`Application.Run` from COM can't assign module-level variables.** It
only invokes named Subs. That's why `gHeadless` and `gTracePath` have
explicit setter Subs (`SetHeadless` / `SetTrace`) the host calls by
name.

**`MsgBox` blocks `Application.Run` from COM.** With Excel hidden, a
MsgBox appears in the hidden VBE window and the COM call never
returns. Every public workflow Sub checks `gHeadless` before showing
a MsgBox; verifier scripts set it via `Application.Run "SetHeadless",
$true` before kicking off work. Cell + StatusBar state stay so
results are still observable when headless.

### 9.3a AGOL Map column + Send-to-AGOL workflow

The Sites table has **two** ArcGIS Online links per row:

1. **NFC Map** (col 19, `COL_NFCMAP` — col 18 before the §7c column
   reorder) — state-aware since §9.3b: MI opens
   the FEMA-hosted NFC/ACUB webmap (`webmap=6a1702b9147243d1a5ee62cd614bc681`
   on `fema.maps.arcgis.com`) centered + markered on the row's coords; the
   previous URL pointed at the MDOT Experience app, whose popup chrome
   blocked the inspector from clicking through to the actual damage
   point, so the FEMA webmap link (same underlying data, no UI wrapper)
   replaced it. IN/WI and any other state get a different URL - see §9.3b.

2. **AGOL Map** (col 28, `COL_AGOLMAP` — col 24 before the §7c column
   reorder) — driven by the `JobAgolMap` cell on Start Here. When the
   user pastes their own AGOL webmap URL there, every Sites row gets a
   hyperlink that opens that webmap centered + markered on the row.
   Empty cell when the URL is blank, so the column is unobtrusive when
   nobody has wired it up.

The URL-stitching formula auto-detects whether the pasted URL already
has a `?` (most AGOL share URLs do, e.g.
`https://www.arcgis.com/apps/mapviewer/index.html?webmap=<id>`) and
uses `&` or `?` accordingly.

To push the *whole site set* into the inspector's AGOL webmap, the
Maps sheet has a **Send Sites to AGOL Map (KML + open webmap)**
button. It writes `RoadReviewer Sites.kml` into the output folder,
opens the AGOL webmap in the default browser, and pops Explorer with
the KML highlighted so it's a single drag-drop into the Map Viewer
window. AGOL ingests the KML as a hosted-feature layer. (V2 could
push directly via the AGOL REST API, but that needs auth and is
out of V1 scope per §3.3.)

### 9.3b State-aware NFC Map column

The NFC Map column formula (`modBuild.SetNfcMapFormula`) and the CSV
export's equivalent resolver (`modExport.NfcMapUrlForRow` - the two
must be kept in sync by hand, there's no shared helper across the
formula-string world and the plain-VBA-string world) both branch on
Setup's `JobState` cell:

- **MI** → `URL_NFC_MAPVIEW`, the curated FEMA-hosted webmap described
  in §9.3a.
- **IN** → `URL_NFC_MAPVIEW_IN`, **WI** → `URL_NFC_MAPVIEW_WI` - side-load
  the state's own live NFC FeatureServer (§4.2a/§4.2b) via the Map
  Viewer's `url=` parameter, using the same `find=`/`marker=`/`level=`
  parameters already proven to work for `URL_FEMAVIEW` (rather than the
  `webmap=`+`center=` combination MI's link uses, which broke once
  before when a `visibleLayers` param was added - see the `URL_NFC_MAPVIEW`
  comment in `modConstants.bas`). No curated combined NFC+ACUB webmap
  exists for IN/WI to point at instead.
- **Anything else** (MN/IL/OH, or a blank State cell) → `URL_FEMAVIEW`,
  the plain pin+zoom with no data layer. A blank State cell is treated
  as MI instead, matching `ClassifyRows`'s default-to-MI behavior - Setup
  pre-fills `JobState` to `"MI"` at build time, so this only matters if
  someone manually clears the cell.

**Not yet click-tested.** The `url=`-sideload URLs were built by
reasoning from a working pattern (`URL_FEMAVIEW`) and the ArcGIS REST
`url=` parameter's documented behavior, not by opening them in a
browser - this repo's cloud sandbox can't drive a JS single-page map
app. Worst case if the layer overlay doesn't render: the link still
opens to the correct pin+zoom, just without the NFC layer on top, so
this degrades gracefully rather than breaking. Confirm the layer
actually renders during the next local smoke pass.

### 9.4 Excel / COM / PowerShell quirks

**The repo's `build\` folder AND the repo root have `Everyone Deny
DeleteSubdirectoriesAndFiles`** (Claude Code sandbox guard). Excel's
SaveAs writes to a temp file in the destination dir, then deletes the
existing destination and renames — the delete step is blocked, which
killed direct `SaveAs` to either path on overwrite. `build\build.ps1`
sidesteps it: SaveAs lands in `%TEMP%\rr-build-<guid>.xlsm` (always
writable), then `Copy-Item -Force` overwrites the OutPath without
needing delete permission on the parent dir. The committed
`RoadReviewer.xlsm` lives at the repo root so inspectors can
double-click straight after `git pull`.

**PowerShell COM late-binding sometimes picks the wrong overload of
`Range.Value`.** Setting a cell value via `$cell.Value = $someDouble`
can throw *"Unable to cast object of type 'System.Double' to type
'System.String'"*. Use `.Value2` with explicit casts:
`$sites.Cells(2, 6).Value2 = [double]$lat`.

**Mark-of-the-Web on a committed `.xlsm`.** OneDrive sometimes adds
the MoW zone after a sync, which makes Office disable macros on
first open. Fix per file: right-click → **Properties** → tick
**Unblock** at the bottom → **OK** → reopen → click **Enable
Content** on the yellow bar.

### 9.5 Blank-WO/DI handling (PR #5)

Setup's WO and DI are **not required**. Every site that emits them
in a filename, folder path, or textbox stamp routes through
`modUtil.JobIds(wo, di, sep, woPrefix, diPrefix)`, which drops the
blank piece and only inserts the separator between two present
pieces:

| Call site | Separator | Prefixes | Both blank → |
|---|---|---|---|
| `FirmetteFileName` | `" "` | `WO` / `DI` | piece omitted |
| `ExportCombinedMapPdf` filename | `" "` | `WO` / `DI` | piece omitted |
| `DefaultOutputFolder` segment | `"-"` | `WO` / `DI` | folder segment skipped |
| `BuildMapTextboxString` first line | `" "` | `WO #` / `DI #` | line skipped entirely |

A row's own WO/DI cells (Sites cols A/B) take precedence over the
Setup values for that row, which is the per-row override path noted
in F14.

### 9.6 Useful trace switches

```vba
' From the VBE Immediate window, to start writing every HTTP call to a file:
SetTrace Environ$("TEMP") & "\rr_trace.txt"
' Then run any workflow. Lines look like:
'   11:43:06 HTTP GET https://...
'   11:43:07   -> 200 (697 bytes)
'   11:43:07 ClassifyOneRow row=2 name=...
' Turn it off:
SetTrace ""
```

`gHeadless` can be flipped the same way:

```vba
SetHeadless True    ' suppress every workflow's MsgBox
SetHeadless False   ' restore default
```

