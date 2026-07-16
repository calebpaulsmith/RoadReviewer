# RoadReviewer

A FEMA Public Assistance road-review toolkit. Replaces and consolidates two
prototype workbooks (`GPS Checker - TN updated 3.5.2026.xlsm` and
`Site Inspector Tool 1.xlsm`, now in `archive/`) into polished, macro-driven
Excel workbooks that staff can actually use without training.

> Historical narrative (prototype autopsy, build increments 1-6, the web
> tool's design story, PR #21-#32 play-by-play) lives in **`docs/HISTORY.md`**
> — consult it for archaeology; nothing there governs the current build.

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

## 2. Existing prototypes — summary

The two prototype workbooks (`GPS Checker - TN updated 3.5.2026.xlsm`,
`Site Inspector Tool 1.xlsm`) now live in `archive/` as read-only references.
Their full autopsy (sheet/column maps, VBA module inventories, findings) is in
`docs/HISTORY.md` §2. The absorbed lessons: one shared Sites table; hyperlink
formulas off lat/lon; Setup named ranges for job-wide values; the FIRMette
submitJob→poll→OutputFile flow; per-row status columns; red-flag rule for
federal-aid classes; and never write to ActiveSheet — hard-bind every sheet.

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
- **F5. All six Region V states wired for road classification.**
  Hard-code each state's own NFC layer(s) for class + road name (MDOT for
  MI, INDOT/IndianaMap for IN, WisDOT for WI — WI queries a local-roads
  layer first and falls back to a state-trunk layer). **MN / IL / OH were
  wired in PR #36** (§4.2c-e): all three publish a bare FHWA 1-7 code,
  Indiana-shaped, no active/retired filter. *(See*
  [§4 Data sources](#4-data-sources) *.)*
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
- **F8. State selector.** The State dropdown offers all six Region V
  states (WI / IN / MI / MN / IL / OH), every one wired for NFC lookup
  since PR #36. A TYPED out-of-region code (e.g. TN) still pops the
  "NFC lookup not yet wired — ACUB still runs" message and continues
  ACUB-only. A BLANK State refuses to classify: every target row gets
  "Failed - no State selected" plus a prompt pointing at the State box
  (blank used to silently mean Michigan — removed in PR #36).
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
sequence**: the Local Road Network layer covers county/municipal/town
roads *and most collectors* (the large majority of points) but encodes
urban/rural into its own class code instead of carrying a bare FHWA
number; the State Trunk Network covers interstates/state highways with a
clean bare FHWA code.

**Query order (updated — PR "WI layer swap", 2026-07-10): local FIRST,
trunk as fallback.** Per user direction, the Local Road Network layer is
queried first because it carries most points; the State Trunk Network is
consulted only when the local layer has no usable class for the point.
This reverses the earlier trunk-first order. See "Query strategy" below
for the stub-detection detail that keeps state highways correct.

**Link change (same PR): WisDOT moved/locked the old local-roads
service.** The previous `WI_Local_Roads_Flood_Damage_Assessment_Snapshot`
now returns `{"error":{"code":499,"message":"Token Required"}}`; the live
public local layer is `Functional_Class_Local_Non_Prod/1` (same schema,
same account, same fields). The state-trunk `FFCL_gdb/3` was unchanged.
Both URLs are now **overridable at runtime** without a rebuild — paste a
replacement into the Sources sheet's "Service URLs" table (Excel named
range `Svc_WI_LOCAL_ROADS`), the web tool's "Data service URLs" panel, or
the notebook's `SERVICE_OVERRIDES` dict (see "Service-URL overrides"
below).

#### Layer `FFCL_gdb/3` — State Trunk Network (fallback; state highways/interstates)

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

#### Layer `Functional_Class_Local_Non_Prod/1` — Local Road Network (queried FIRST; local roads + most collectors)

- URL: `https://services5.arcgis.com/0pgGLzT0Nh7FVjon/arcgis/rest/services/Functional_Class_Local_Non_Prod/FeatureServer/1`
  (name "Functional Class - Wisconsin Local Road Network"; replaced the
  token-gated `WI_Local_Roads_Flood_Damage_Assessment_Snapshot/1`,
  confirmed live 2026-07-10).
- Type: `Feature Layer` (esriGeometryPolyline), Web Mercator (`inSR=4326`
  reprojects). Same WISLR-derived ~97-field schema as WisDOT's
  general-purpose `WISLR_Functional_Road_Classification` hosted layer,
  filtered to the local-road network. **State highways also appear here
  but with a null (or `0`) `FNCT_CLS_CTGY_TYCD` — an unclassified "stub."**
  The classifier skips those stubs for class/name AND treats their
  presence as the trigger to also query the State Trunk layer (so a point
  on a state highway is still classified). Confirmed live 2026-07-10: at
  Milwaukee the layer returns `E Wisconsin Ave` (code 60) alongside a
  null-class `'18'` stub; at STH 52 Rhinelander it returns only `'52'`
  stubs (no real local class → fall through to trunk).
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

#### Query strategy (Wisconsin) — local-first (PR "WI layer swap")

1. Point-intersect (then configured-buffer fallback) against
   `Functional_Class_Local_Non_Prod/1` **first** (one query returns both
   `FNCT_CLS_CTGY_TYCD` and `ST_LABL_NM`, halving the hit count on this
   layer). For each returned feature: if `FNCT_CLS_CTGY_TYCD` decodes
   (via `WisconsinLocalCategoryToFhwa()`) to a real FHWA class ≥ 1, keep
   the segment + name; if it is null or decodes to `0` (a state-highway
   stub), skip it and set `sawStub`.
2. Query `FFCL_gdb/3` (State Trunk, bare `FED_FC_CD`) **only when** the
   local layer returned no real class **OR** `sawStub` is set (a state
   highway is present that only the trunk layer classifies). Merge its
   segments/names in with the local ones.
3. The **closest** road across whatever both layers returned drives the
   verdict (PR #24 model). Net effect: a point genuinely on a local road
   near a state highway now reads as the local road with a yellow
   "Nearby FHWA road" flag, rather than being forced to the highway's
   class (e.g. 45.4711,-89.7345 = W Wisconsin Ave 26 ft vs STH 86 199 ft).
4. ACUB (urban/rural) runs exactly as it does for every other state —
   independently of whichever WI layer answered the class query.

Rationale for local-first (per user, 2026-07-10): most damage points are
on local roads / collectors carried by the local layer, so querying it
first means one query for the common case; the trunk layer is a low-risk
fallback (it is the layer WisDOT did *not* lock down). The stub trigger
keeps the single-query optimization from ever misclassifying a state
highway that happens to sit near a local road.

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

### 4.2c Minnesota road classification (wired PR #36)

Source: MnDOT's public ArcGIS Server (`dotapp9.dot.state.mn.us`), the same
Functional Class layer MnDOT's EMMA app (`APP_MN`) displays. Schema
confirmed live 2026-07-15 via `?f=pjson` + `returnDistinctValues`.

- **Layer:** `https://dotapp9.dot.state.mn.us/egis12/rest/services/BASEMAP/mndot_commonlayers2/MapServer/11`
  ("Functional Class", esriGeometryPolyline, native WKID 26915 / NAD83 UTM
  15N — hand it WGS84 with `inSR=4326`). Classic MapServer, `capabilities:
  Query,Map,Data`.
- **Class field:** `FUNCTIONAL_CLASS` (esriFieldTypeInteger) — **bare FHWA
  1-7**, confirmed by distinct values; `FUNCTIONAL_CLASS_DESC` carries the
  matching standard label ("Principal Arterial - Interstate" … "Local").
- **No active/retired filter needed** (`where=1=1`).
- **No road-name field** — `ROUTE_ID` is an LRS key. Census TIGER backfills
  all Minnesota street names.
- **Renderer quirk (web figures):** MnDOT's published uniqueValue renderer
  keys off `FUNCTIONAL_CLASS_DESC` (the label string), not the numeric
  code — the web tool's frame query fetches both fields.
- **Test coordinates (live-verified 2026-07-15):**

  | # | Expected | lat | lon | Notes |
  |---|---|---|---|---|
  | 1 | Federal aid - Urban Minor Arterial | `44.9531` | `-93.1668` | Snelling Ave, St Paul (class 4 @ 36 ft closest); ACUB `Minneapolis--St. Paul, MN` |
  | 2 | Non-federal aid - Urban Local | `44.9260` | `-93.2570` | Minneapolis residential, all class 7 |
  | 3 | Non-federal aid - Rural Local | `45.822764` | `-95.222414` | rural Douglas County, class 7 @ 0 ft, outside every ACUB |

### 4.2d Illinois road classification (wired PR #36)

Source: IDOT's public ArcGIS Server (`gis1.dot.illinois.gov`), the layer
behind the Getting Around Illinois RFC viewer (`APP_IL`). Confirmed live
2026-07-15.

- **Layer:** `https://gis1.dot.illinois.gov/arcgis/rest/services/AdministrativeData/FunctionalClass/MapServer/0`
  ("Functional Class", polyline, WKID 102113/3785 Web Mercator).
- **Class field:** `FC` — a **STRING** `"1"`-`"7"` with a published
  coded-value domain matching FHWA (1 Interstate (PAS) … 7 Local Road or
  Street). Read with `isStringClass=True` in `AddClassSegs`, like WI's
  `FED_FC_CD`.
- **No active/retired filter needed.**
- **No street names** — `LABEL_1`/`KEY_RT_*` are route-system inventory
  codes ("FAU 1422", "MUN 4022H"), not names. TIGER backfills.
- **Test coordinates (live-verified 2026-07-15):**

  | # | Expected | lat | lon | Notes |
  |---|---|---|---|---|
  | 1 | Federal aid - Urban Other Principal Arterial | `41.9020` | `-87.6870` | Western Ave, Chicago (class 3 @ 2 ft); ACUB `Chicago, IL--IN` |
  | 2 | Non-federal aid - Urban Local | `41.9430` | `-87.7010` | Chicago residential, single class-7 |
  | 3 | Non-federal aid - Rural Local | `40.165157` | `-89.434236` | rural Logan County, class 7 @ 0 ft |

### 4.2e Ohio road classification (wired PR #36)

Source: ODOT's TIMS public ArcGIS Server (`tims.dot.state.oh.us`), the
Functional Class layer TIMS (`APP_OH`) displays. Confirmed live 2026-07-15.

- **Layer:** `https://tims.dot.state.oh.us/ags/rest/services/Roadway_Information/Functional_Class/MapServer/0`
  ("Functional Class", polyline, WKID 102100/3857 Web Mercator).
- **Class field:** `FUNCTION_CLASS_CD` (esriFieldTypeInteger) — **bare FHWA
  1-7**, confirmed by distinct values.
- **No active/retired filter needed.**
- **Route names on the same layer:** `ROUTE_TYPE` + `ROUTE_NBR` read
  cleanly for the numbered systems — `"US"/"00023"` → "US 23",
  `"SR"/"00161"` → "SR 161", `"IR"` = interstate. County/township/municipal
  codes ("MR 00923") are cryptic and skipped (`RoadNameFromBlock`'s
  `OH_ROUTE` mode); TIGER fills local street names.
- **Test coordinates (live-verified 2026-07-15):**

  | # | Expected | lat | lon | Notes |
  |---|---|---|---|---|
  | 1 | Federal aid - Urban Other Principal Arterial | `40.0150` | `-82.9990` | N High St / US 23, Columbus (class 3 @ 25 ft; Road Name shows "US 23 (25 ft)"); ACUB `Columbus, OH` |
  | 2 | Non-federal aid - Urban Local | `40.0855` | `-83.0170` | Columbus residential, single class-7 |
  | 3 | Non-federal aid - Rural Local | `40.320352` | `-83.302785` | rural Marion County, class 7 @ 3 ft |

#### UA / header quirks (all three)

None — all three servers answered plain GETs without a browser User-Agent
(RoadReviewer sends its browser UA regardless, harmless). Note: INDOT's
`gisdata.in.gov` showed intermittent 500s / empty responses on 2026-07-15
during this pass (recovered within minutes each time) — unrelated to these
three, but a reminder that per-row Failed + Re-run Failed Rows is the
designed recovery for state-server flakiness.

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

- **Auto-capture basemap for MapPages — SHIPPED (PR #35, §7e).** "Fetch
  Imagery" on the Map Pages ribbon downloads a site-centered Esri World
  Imagery aerial per page (one anonymous GET against the MapServer
  `/export` endpoint) and places it automatically; the imagery source is
  user-swappable to any ArcGIS MapServer via the Map Pages "Imagery URL"
  cell. Remaining risk from the original list: reachability of
  `services.arcgisonline.com` from a hardened govt laptop is untested.
- **Route-optimization integration.** The user has a separate route
  planner. Easiest hand-off: export the Sites table as a CSV / KML / GPX
  that the planner already accepts. Avoid tight coupling for now.

- **Rapid site-by-site review on the authoritative map (idea, PR #22).**
  The tool already exports a KML of all sites (`ExportSitesToKML`) and can
  drop it onto the user's AGOL webmap (`SendSitesToAgolMap`). Now that the
  official public functional-class app URL is known for MI/IN (and MN/IL/OH
  — see §7c / `modSources`), two follow-ons are feasible: (a) a per-row
  deep-link that opens *that state's official app* centered on the site
  (today the NFC Map column opens the FEMA Map Viewer webmap, not the DOT's
  own Experience app — the Experience Builder coordinate deep-link was
  rolled back once for mis-navigating, PR #18, so this needs care); and
  (b) publishing the Sites KML as a hosted feature layer and building an
  ArcGIS **Instant App** (e.g. Attachment Viewer / Media Map / Nearby) to
  step through each site over the authoritative class + ACUB layers. Both
  need an AGOL org + auth, so they're out of the no-auth V1 scope, but the
  verified per-state app/REST URLs are the missing piece and are now on
  file.
- **AGOL Experience Builder front-end.** For features that are awkward in
  Excel — drawing a polygon search area, side-by-side historical imagery,
  collaborative editing — we can publish the Sites table as a hosted
  feature service and build an Experience app on AGOL. The Excel
  workbook stays the *system of record* (the inspector still owns their
  spreadsheet); the AGOL view is a read-only or write-on-top layer.
  This sidesteps any limitation of "only public maps" because AGOL is
  already licensed for the user's organisation.
- **Historical Google Earth / Nearmap links** for pre-disaster imagery.
- **State expansion — DONE for Region V (PR #36).** All six states are
  wired (§4.2-§4.2e); no normalisation function was needed for MN/IL/OH
  (all bare FHWA 1-7). ACUB stays nationwide. Any future out-of-region
  state follows the same recipe: probe `?f=pjson` + `returnDistinctValues`,
  add a `QueryStateRoads` case + REST/Svc_ key, live-verify 3 coords.
- **NHS column.** Surface MDOT's NHS layer (and the equivalent in other
  states) so the inspector can see federal-aid-system status alongside
  functional class.

---

## 7. Repository layout

```
/                                  this README — CLAUDE.md
.github/workflows/pages.yml       deploys web/ to GitHub Pages on push to main (§7b)
archive/                                the two prototypes + superseded workbook builds (reference only)
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
  modMapFetch.bas                       Fetch Imagery — auto-download a site-centered aerial
                                          per map page (Esri World Imagery /export or a pasted
                                          ArcGIS MapServer), yellow pushpin, attribution, re-run
                                          failed (§7e)
  modPdf.bas                            direct PDF writer for the Location Map export (§7g) —
                                          JPEG pages + vector stamp/pin/attribution; bypasses
                                          Excel's (machine-dependently broken) print renderer
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
  verify-imagery.ps1                    §7e — Fetch Imagery end-to-end: 3 sites, failure path,
                                          re-run-failed, imagery-URL override, PDF checked with
                                          PyMuPDF (page count, attribution text, pixels, pin)
  verify-screenshot-pdf.ps1             §7g — manual screenshot flow + exact PDF block geometry
                                          (760x568 on every page), incl. a sheet-sabotage leg
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
  HISTORY.md                            historical narrative moved out of this file (prototype
                                          autopsy, increments 1-6, web design story, PR #21-#32)
  Region V Test Coordinates.xlsx        the 19 live-verified test points, all six states (§7f)
notebooks/                              AGOL notebook port of rr-core (jupytext-paired .py/.ipynb)
```

The two prototype `.xlsm` files are kept in `archive/` as references. The V1
deliverable is a new workbook so we don't drag along the VBA-stomping
warning from the Site Inspector Tool. The workbook **is committed at the
repo root** so it's accessible to inspectors who clone the repo, and is
rebuilt on the user's Windows laptop from `src/` (see
`docs/build-and-import.md`) — the cloud build environment has no Excel,
so changes to `src/` need a local rebuild + commit before the .xlsm
reflects them.

---

## 7a. Implementation status (V1)

Everything in the §3 V1 scope is BUILT and covered by the §9.2 verifier
suite: all six Region V states classify live (PR #36), Fetch Imagery is the
map hero (PR #35, §7e), FIRMettes/Map Pages/KML/CSV/GeoJSON all pass their
verifiers on the committed workbooks. The increment-by-increment build
narrative (increments 1-6, per-capability status table, the bugs each pass
found) is in `docs/HISTORY.md` §7a; the recurring VBA/COM traps distilled
from it live in §9.3/§9.4 below.

---

## 7b. FHWA Road Checker — public web prototype (summary)

`web/index.html` is RoadReviewer's Workflow 1 as a public static page:
client-side only (no backend/accounts/analytics; the browser queries the same
public services), paste-and-see UX, PR #24 verdict model ported 1:1, PDF
report + FIRMette ZIP + site-by-site review. rr-core (the `<script id="rr-core">`
block) is a hand-port of modClassify/modConstants — keep the two in sync by
hand; `build/verify-web-core.mjs` runs the exact shipped block against the
live services. Full design narrative + verification history:
`docs/HISTORY.md` §7b; privacy/hosting: `web/README.md`; per-state citations:
`web/sources.html`. Load-bearing gotchas:

- Experience Builder coordinate deep-links MIS-NAVIGATE (PR #17/#18): link
  official state apps at their canonical ROOT; pinned-location duty sits on
  the FEMA Map Viewer links.
- AGOL-hosted FeatureServers (IN/WI/ACUB) are Query-only — no /export or
  /legend. Figures draw each layer's published drawingInfo.renderer
  client-side (INDOT publishes single-symbol → FHWA-palette substitution,
  disclosed). WI trunk renderer keys off FC_CD, not FED_FC_CD; MnDOT's keys
  off FUNCTIONAL_CLASS_DESC.
- jsPDF must be created `compress: true` (else ~5 MB/figure); basemap tiles
  load `crossOrigin="anonymous"` against Esri's CORS tile service.
- Leaflet map is created `zoomAnimation: false` (in-flight zoom animations
  swallow the next setView).
- GitHub Pages: "Re-run all jobs" fails with duplicate github-pages
  artifacts — recover with a FRESH run, or re-run only the deploy job.

---

## 7c. Two-product split (2026-07-05) — summary

Two products, one `src/` tree, same three-sheet shape (§7d has the CURRENT
sheet roles). Do not relitigate the split. PR-by-PR narrative (#21-#32,
including now-OBSOLETE column orders) is in `docs/HISTORY.md` §7c.

| | RoadReviewer.xlsm (Standard) | Site Inspector Review Tool.xlsm (Inspector) |
|---|---|---|
| Audience | PDMGs, state/local partners, reviewers | FEMA PA site inspectors |
| Extra inputs | — | WO, DI, Disaster, Applicant |
| Extra buttons | — | FIRMettes, Map Pages workflow (hero) |
| Column hiding | inspector-only cols hidden (WO/DI/FIRMette+Map Status) | photo/NFC links hidden until Check Roads |

Product identity: hidden defined name `RR_Product` baked at build time
(`modUtil.SetProduct`, called by build.ps1); runtime branches via
`ProductIsInspector()` / `ProductTitle()` / `StartSheetName()`; a missing
name defaults to Inspector (superset). `COL_*` constants are SHARED — the
standard product hides inspector-only columns rather than having its own
column map, which keeps modClassify/modMaps/modExport product-agnostic. The
in-Excel "Repair Layout" button rebuilds the same product. Check Roads =
auto-geocode (Address rows without coords, never overwrites) + classify;
Re-run Failed Rows retries anything whose Federal Aid Status starts
"Failed - ". Current column order: `src/modConstants.bas` (§7f). Workbooks
filled under pre-split layouts do NOT migrate; start new work from the
current builds.

---

## 7d. Map Pages hero redesign + "Tools and Exports" (2026-07-14) — CURRENT shape

A full UI restructure of the inspector product, per user direction across one
long session. **This supersedes the §7c sheet layout for the inspector.** The
prior design is archived: tag `start-here-hero` (commit `74b1d43`) + byte-exact
workbooks in `archive/start-here-hero/`.

### Sheet roles per product

| | Inspector (`Site Inspector Review Tool.xlsm`) | Standard (`RoadReviewer.xlsm`) |
|---|---|---|
| Landing (tab 1) | **Map Pages** (visible) | **Start Here** (visible hub, unchanged role) |
| Tab 2 | Sites | Sites |
| Hidden | **"Tools and Exports"** (the old Start Here, renamed; reached via the grey "Exports & other tools →" button on Map Pages, `GoToOtherTools`; "← Back to Map Pages" = `GoToMapPages`) and Sources | **Map Pages** (opt-in: any map action calls `modMaps.ShowMapPages` and reveals it with FULL functionality) |

The hub sheet is named per product — `modUtil.StartSheetName()` returns
`SH_START` ("Start Here", standard) or `SH_TOOLS` ("Tools and Exports",
inspector); every hub lookup goes through it. The map tab is `SH_MAPPAGES` =
**"Map Pages"** (renamed from "MapPages" — verifiers with hardcoded sheet-name
literals broke once already).

### Map Pages is a PERMANENT sheet

Built by `BuildWorkbook` → `modMaps.EnsureMapPagesSheet` (it used to be created
and DELETED by `PrepareMapPages`; that lifecycle is gone). `PrepareMapPages`
rebuilds only the page blocks (`ClearMapPages`) and preserves the header band +
job values — safe to re-run. The header band (rows 1..`MAP_HEADER_ROWS`, short
rows; pages start at `MAP_FIRST_PAGE_ROW`; PrintArea starts there so the band
never prints) holds:

- the ONE-CLICK hero (2026-07-16, superseding PR #35's 3-step ribbon per
  user direction "make the three buttons a single button"): **Create Combined
  Map Pages PDF** = `CreateMapPagesPdf`, which chains `PreparePagesCore` →
  `FetchImageryCore` → `ExportMapPdfCore` with a single summary MsgBox.
  **Download FIRMettes** sits right beside it as the other primary
  deliverable. Everything else lives behind a collapsed **"Advanced
  options ▸"** ghost toggle (`ToggleMapAdvanced`; state in the toggle shape's
  AlternativeText, respected by `SetMapEditControlsVisible` when it restores
  controls after an export): the individual steps (1 Prepare Pages / 2 Fetch
  Imagery / 3 Export PDF), ghost "↻ Re-run failed imagery" / "↻ Re-stamp
  pages" / "↻ Re-run failed FIRMettes", **Create Individual Map Pages PDFs**
  (`CreateIndividualMapPagePdfs` — one PDF per site named exactly like the
  FIRMette but "… Location Map.pdf"; direct writer per page,
  `ExportOnePagePrint` fallback), the **Exports & other tools →** door
  (`GoToOtherTools`, inspector only — moved here from the top-right), and the
  labeled **manual alternative** (Export to KML + Insert Images — the Google
  Earth screenshot flow). Advanced shapes are named `MapCtrl_Adv*` and ship
  hidden;
- the job block: WO / DI / Disaster (bare number, e.g. 4882) / **State**
  (inspector only; drives classification AND the filename tag) / Applicant /
  Output Folder (optional, + Browse) / **Imagery URL (optional, PR #35)** —
  these carry the `JobWO`/`JobDI`/`JobDisaster`/
  `JobState`/`JobApplicant`/`JobOutputFolder`/`JobImagerySvc` named ranges on the inspector
  (standard keeps State/Output Folder/AGOL/buffer on Start Here; its Map Pages
  shows a read-only Output Folder mirror so the name isn't defined twice).
  Band constants moved for the new layout: `MAP_HEADER_ROWS` 20→22,
  `MAP_JOB_FIRST_ROW` 11→14;
- to the RIGHT of the job boxes, a **formula-driven stamp preview** (cols I:M,
  `StampPreviewFormula`) — a live Excel-formula replica of
  `BuildMapTextboxString` driven by the first Sites data row + the job named
  ranges (blank lines skipped via a leading-`CHAR(10)` strip), so the
  inspector sees each page's stamp as they type — with the **filename
  preview** (`="File name:  "&FirmettePreview()`, volatile) directly under it.

### Shared file-name convention (modUtil)

`JobFileStem()` = `"WO123 DI5 - DR4882MI"` → used by **every** export
(Location Map PDF, per-site FIRMettes, KML, CSV, GeoJSON), suffixed per type
(`... - Sites.kml` etc.). `DisasterTag()` composes the user's convention —
**no separators since PR #37: `DR4882IN`, not `DR-4882-IN`** — bare digits
get `DR`, typed prefixes are kept with any hyphens/spaces stripped, State
appended when set (unless already typed). All-blank job info falls back to a
`yyyy-mm-dd HHmm` stamp so repeated exports never overwrite.

### State handling

State ships **blank by default** (both products). The two state-dependent NFC
link columns (13 "NFC Layer (Map Viewer)", 15 "State NFC App") show — on the
FIRST data row only, other rows blank — a working `HYPERLINK("#JobState","Set
State →")` instead of silently defaulting to Michigan. Since PR #36
`modClassify` refuses to run with a blank State (every target row gets
"Failed - no State selected" + a prompt naming the State box) — blank no
longer silently means Michigan. The State dropdown lists all six wired states
(`STATE_LIST = "WI,IN,MI,MN,IL,OH"`); the ACUB-only path for typed
out-of-region states still works.

### Hard-won operational rules (cost real corruption TWICE)

- **Verifiers must NEVER open the committed workbook read-write. Period.**
  `Close($false)` is NOT enough — OneDrive AutoSave persists macro side
  effects the moment they happen. This bit twice: first the verifiers that
  ended `$wb.Save()`, then AGAIN on 2026-07-16 when verify-firmette-maps
  (temp-copy rule not yet applied to it) deleted the permanent Map Pages
  sheet as "cleanup" and AutoSave synced `JobWO`→`#REF!#REF!` into the
  committed inspector file. Every verifier that runs macros now COPIES the
  workbook to %TEMP% first; verify-skeleton (inspection only) opens READONLY.
  A corrupted committed workbook is repaired by a plain rebuild — the .xlsm
  is fully generated from `src/`.
- **Never delete the Map Pages sheet** (in a verifier or anywhere): it is
  permanent (§7d) and carries the JobWO/JobDI/... named ranges.

## 7e. Fetch Imagery — automatic aerial download for Map Pages (PR #35, 2026-07-15)

Map Pages no longer needs manual screenshots as its primary flow. The ribbon
hero was **1 Prepare Pages → 2 Fetch Imagery → 3 Export PDF** — since
2026-07-16 those three are chained behind the single **Create Map Pages PDF**
button and live individually under "Advanced options" (§7d/§7g); the Google
Earth screenshot path (Export to KML + Insert Images) is preserved there as
the labeled "Manual alternative" (user direction: keep it as an option, not
the hero). The **PR #34 workbooks are archived byte-exact** in
`archive/pr34-manual-screenshots/`.

**How it works (`modMapFetch.bas`):** for every map page that references a
Sites row, one anonymous GET against Esri World Imagery's MapServer
`/export` endpoint (`REST_WORLD_IMAGERY`, proved live 2026-07-14/15: HTTP
200, image/png) returns a rendered aerial for a Web-Mercator bbox centered
on the site — half-width `MAP_IMG_HALFWIDTH_M` (600 m), half-height scaled
to the page-block aspect (760:568), pixels at 2× (1520×1136) so the crop
step in `PlaceImageOnPage` is a no-op. The PNG lands in `%TEMP%`, is placed
through the existing `modMapImage` pipeline (crop-to-cover + 1 pt inset
kept), and is also copied to `<output folder>\maps\Site_<n>.png` so the
manual Insert-Images flow interoperates. Because the site is always the bbox
center, a small red **pin shape** is dropped at the geometric center of each
page, plus the Esri-required **attribution** textbox bottom-left
(`MAP_PIN_PREFIX` / `MAP_ATTR_PREFIX` — deliberately NOT the control
prefixes, so `SetMapEditControlsVisible` never hides them: they must print).
Per-row progress goes to the Map Status column with `STATUS_FAILED_PREFIX`,
and **Re-run failed imagery** retries only failed rows (FIRMette model). A
row whose coordinates went bad after Prepare is *processed and marked
Failed* (not skipped), so failures are visible and re-runnable.

**User-swappable imagery source.** The Map Pages job block has an
**"Imagery URL (optional)"** cell (`JobImagerySvc`, both products): blank =
Esri World Imagery (also overridable repo-wide via `Svc_WORLD_IMAGERY` on
Sources); paste any **ArcGIS MapServer** URL to fetch from it instead
(`ImageryServiceBase` normalizes a browser-copied URL — strips query string,
trailing slashes, trailing `/export`). Only MapServers work: the `/export`
operation doesn't exist on Query-only FeatureServers (same constraint as the
web PDF figures, §7b). With a custom source the attribution line switches to
"Imagery: <host>".

**Traps hit during this pass (all fixed, all verified):**

1. **`ShowMapPages` before the early-exits = the §7d AutoSave trap.**
   `compile-check.ps1` no-op-runs `ReRunFailedImagery` against the COMMITTED
   workbook; an unconditional `ShowMapPages` at the top of the sub unhid the
   standard product's Map Pages and OneDrive AutoSave persisted it (caught by
   verify-skeleton). Fix: reveal only after the "any pages?" check, and
   compile-check now opens READONLY.
2. **`modMapImage.ClearPlaceholderText` still used the pre-header-band row
   formula** (`pageIdx*ROWS+1`), so on pages 3+ it cleared cells INSIDE the
   header band (job-field labels) instead of the page placeholder. Latent
   since the §7d band was added; fixed to the `MAP_FIRST_PAGE_ROW` offset.
3. **`build.ps1` imports an explicit module list**, not a glob — a new .bas
   must be added there or it silently doesn't ship (compile-check catches it).
4. **The 1-site "single cell for the print area" modal.** Each map page is
   one big MERGED cell, so with exactly one site the print area is a single
   merged cell — and Excel pops "You've selected a single cell for the print
   area" on the `PageSetup.PrintArea` assignment whenever alerts are on
   (`PrepareMapPages` re-enables `DisplayAlerts` before Export runs). It hung
   every headless 1-site export (an invisible `NUIDialog` — found by
   enumerating the hidden Excel's windows via user32) and hit the user
   interactively too. OK is the right answer, so `SetMapPrintArea` now wraps
   the assignment in its own `DisplayAlerts = False` save/restore. Multi-site
   runs never showed it, which is why verify-imagery (3 sites) passed while
   verify-firmette-maps (1 site) hung.

**Layout constants moved:** `MAP_HEADER_ROWS` 18→20, `MAP_JOB_FIRST_ROW`
8→11 (the manual-alternative line sits between the ribbon and the job
block). `modHttp.HttpDownloadPdf` was generalized into `HttpDownloadBinary`
(accept header + expected Content-Type as params; PDF wrapper kept).

**Verification:** `build\verify-imagery.ps1` runs end-to-end on a %TEMP%
copy of the inspector build: 3 live sites → Prepare → one row sabotaged →
Fetch (2 placed, 1 Failed, batch continues) → sentinel + re-run-failed (only
the failed row retried) → imagery-URL override re-fetch (World_Street_Map,
attribution switches) → Export PDF → PyMuPDF asserts page count == 3,
attribution + stamp text on every page, ~90% non-white imagery coverage, and
yellow pushpin pixels just above page center (per §9.8: rendered pixels,
never `get_image_rects`). Known risk (stated in the commit): reachability of
`services.arcgisonline.com` from a hardened FEMA laptop is untested.

## 7f. UX pass: naming, output folder, columns & links (PR #37, 2026-07-15)

All per user direction, both products. Everything below is asserted by the
updated verifiers (skeleton labels/hiding, classify/rerun column indices,
DRTEST filenames) and the whole §9.2 suite passes on the rebuilt workbooks.

1. **Disaster naming**: `DR4882IN`, no hyphens, everywhere the disaster
   appears (see §7d). `DisasterTag` strips typed separators too.
2. **Output folder default**: `<workbook dir>\RR Output\` / `\SI Tool
   Output\` + the OneDrive https→local mapping + `SurfaceFolder` after
   every export (full detail in §8 resolved #9). `SetOutputFolderDefault`'s
   display formula (standard Start Here) appends `RR Output\` so the shown
   default matches. SurfaceFolder is headless-gated, so it's exercised
   only by a human click-test.
3. **AGOL column defaults to the FEMA pin**: blank `JobAgolMap` → every
   row links the FEMA-hosted Map Viewer as "FEMA AGOL Map Viewer"; a
   pasted webmap URL takes the column over as before. The separate FEMA
   Viewer column (27) ships hidden in BOTH products, as do Google Maps
   (23) and Bing (25); Street View stays visible; Google Earth keeps its
   split (inspector hidden). CSV resolver matches.
4. **Reviewer block reordered** — verdicts lead: Federal Aid Status (16/P),
   Review Reason (17/Q), then FHWA Class, Urban/Rural, ACUB Name, Road
   Name, Street Name (18-22). `COL_REVIEWER_FIRST/LAST` still span 16-22,
   so tint/CF/hide logic was untouched; only constants + verifier indices
   moved. **Check Roads reveals the block BEFORE its loop** so the user
   watches verdicts appear row by row.
5. **Map-link columns sit together** (13 "NFC Layer (Map Viewer)", 14
   "State NFC App"), AGOL moved to 15. Labels are descriptive, not "Open":
   col 13 = "Review NFC AGOL Layer", col 14 = "Review State NFC Layer" —
   except **Wisconsin**, whose two links split its two layers: col 13 =
   "Review Local Roads Layer" (local-roads side-load,
   `URL_NFC_MAPVIEW_WI_LOCAL`) and col 14 = "Review State Trunk Hwy Layer"
   (the FFCL trunk side-load, coords included, replacing WisDOT's static
   -PDF app link there).
6. **Excel HYPERLINK 255-char trap** (cost a build round): the
   percent-encoded WI local-roads side-load link is ~261 chars and
   HYPERLINK() returns #VALUE! past 255. The `url=` value is deliberately
   left unencoded (~239 chars, legal — no `&` inside). Watch this on any
   future side-load of a long layer path.
7. **`docs/Region V Test Coordinates.xlsx`** — the 19 live-verified test
   points for all six states (coords + expected verdict/class/ACUB +
   sources), for hand smoke-testing without digging through CLAUDE.md.

## 7g. Direct PDF export + one-click Map Pages (2026-07-16)

Driven by the user's report that a real Location Map PDF had the screenshots
"outside the print area again". Four changes shipped together; the first is
the important one.

### 1. The Location Map PDF is now WRITTEN DIRECTLY (`modPdf.bas`) — Excel's print renderer is broken on real machines

**The discovery (measure output with PyMuPDF, never trust "looks close").**
On the dev laptop, `ExportAsFixedFormat` renders the whole page **vertically
stretched ~7%** (and ~±2% horizontally) while paginating with nominal sizes.
Proven minimal: a bare 100×100 pt square on a *fresh workbook* prints as
97.9×105.1 pt; a 405-pt block of 17-pt rows prints 434.9 pt tall. It affects
cells AND shapes, Brother driver AND Microsoft Print to PDF, interactive AND
headless COM. So every map image printed ~760×606 instead of 760×568,
bleeding to the page edge — the §9.8 work calibrated *pagination* (which is
correct) but the *render* scale was invisible until measured. Two secondary
bugs found on the way and also fixed, though neither was sufficient alone:
`PictureFormat.Crop` metadata is IGNORED by the PDF exporter (the full
uncropped bitmap prints — so crops are now baked into the pixels via WIA in
`PlaceImageOnPage`/`CropImageFileToAspect`), and `xlMoveAndSize` anchoring
let row-height drift stretch pictures (now `xlMove` + an export-time
re-snap: `NormalizeMapLayoutForPrint` → `SnapShapesToPages`).

**The fix:** `ExportMapPdfCore` first calls `modPdf.BuildMapPdfDirect`,
which writes the PDF byte-by-byte: PDF 1.4, one JPEG per page (PNG→JPEG via
WIA, quality 88; JPEG passes through as `/DCTDecode`) cover-cropped into the
760×568 block by a clip path, the stamp + attribution as REAL vector text
(base-14 Helvetica/Helvetica-Bold, WinAnsi — searchable, crisper than
printing), and the site pushpin as vector art. No printer, no page breaks,
no driver margins, no DPI — identical output on every machine.
`verify-screenshot-pdf.ps1` asserts the imagery lands at exactly
(17,23)-(776,589) on every page. The §9.8 print pipeline survives as the
FALLBACK (WIA unavailable, or a page's image source file missing) — the
image source path is remembered in each picture's AlternativeText
(`TagImageSource`; the fetch flow points it at the durable `maps\` copy).
Pages with no image (placeholder state) render as white + stamp.

### 2. One-click hero + Advanced options (§7d has the layout)

"Create Map Pages PDF" (`CreateMapPagesPdf`) chains
`PreparePagesCore → FetchImageryCore → ExportMapPdfCore` (the old public
subs are thin wrappers over these cores and remain as the Advanced-options
step buttons). Everything non-primary is collapsed behind "Advanced
options ▸" per user direction — the band shows exactly two green buttons
(Create Map Pages PDF, Download FIRMettes) plus the job boxes.

### 3. Yellow Google Earth-style pushpin

`modMapFetch.AddSitePin` now builds a grouped vector pushpin (grey needle +
yellow ball + glint) whose **tip** sits exactly on the site; the old red dot
is gone (user direction). Convention: the pin group's bottom-center IS the
site point — `SnapShapesToPages` and `modPdf.PushpinOps` both rely on it.

### 4. Sites photo links (per user): Street View AND Google Earth visible in BOTH products

Google Maps (23), Bing (25), FEMA Viewer (27) stay hidden everywhere; Google
Earth (26) lost its inspector-hidden split.

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
9. **Output folder default — superseded PR #37.** The default is now a
   product subfolder NEXT TO the workbook: `<workbook dir>\RR Output\`
   (standard) or `<workbook dir>\SI Tool Output\` (inspector), created
   on demand; an explicit Output Folder value still wins. A
   OneDrive-synced workbook reports an `https://` path — the old code
   silently fell through to the §8.9 OneDrive probe (the "outputs went
   to Desktop\Scripts\RoadReviewer" bug the user hit); `modMaps.
   OneDriveLocalFolder` now maps the URL back to the locally synced
   path via the `OneDrive*` env vars (longest-matching URL tail whose
   folder actually contains the workbook file). The §8.9 probe remains
   only for a never-saved or unmappable workbook. Every export also
   ends with `modUtil.SurfaceFolder`: the output folder is opened in
   Explorer, or its already-open window is un-minimized and raised.

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

The `url=`-sideload Map Viewer links (the "Review ... Layer" columns,
§9.3b/§7f) remain **unverified in a browser** for every state: built and
reasoned through from the Map Viewer URL parameters already proven for
`URL_FEMAVIEW`, but the sideload rendering itself hasn't been click-tested.
They degrade gracefully (correct pin+zoom even if the layer overlay doesn't
render). `SurfaceFolder`'s raise-existing-Explorer-window behavior (§7f) is
likewise headless-untestable and needs a human click-test.

---

## 9. Developer notes — rebuild workflow and VBA gotchas

Reference material for whoever is editing `src/` next. Lookup-style; the
narrative version of how these bugs got found is in `docs/HISTORY.md` §7a.

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
| `verify-classify.ps1` | §5.4 — Check Roads against the §4.2/§4.2a/§4.2b/§4.2c-e test coords for all SIX states + an address-only auto-geocode row | standard (or either) | MDOT + INDOT + WisDOT + MnDOT + IDOT + ODOT + NTAD + Census |
| `verify-rerun-and-state.ps1` | §5.7 + §5.8 — blank State refuses ("Failed - no State selected"); typed TN gates NFC; MN classifies for real; ReRunFailedRows only retries `Failed - ` rows | standard (or either) | MDOT + MnDOT + NTAD |
| `verify-firmette-maps.ps1` | Inspector workflow 3 — DownloadFirmettes + PrepareMapPages + ExportCombinedMapPdf on one Kalamazoo site | inspector | FEMA Print FIRMette GP |
| `verify-blank-wodi.ps1` | PR #5 — empty WO/DI produces clean filenames + stamps (no dangling `WO `, ` DI`, or `WO #` line) | inspector | FEMA GP |
| `verify-screenshot-pdf.ps1` | §7g — the manual screenshot flow end-to-end with 6 synthetic GE-aspect images (Prepare → Insert → Export), PyMuPDF asserts the imagery fills EXACTLY the 760×568 block on every page and the right image is on the right page; then sabotages the sheet geometry (rows 142→152, a picture displaced/stretched) and asserts the re-export is still exact | inspector | no |
| `verify-imagery.ps1` | §7e — Fetch Imagery end-to-end: Prepare, failure path, re-run-failed-only, imagery-URL override, Export PDF, PyMuPDF pixel/text checks. Copies the workbook to %TEMP% itself. | inspector (either works) | Esri World Imagery export |

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

**A single-line `If` can use `Else` but NOT `ElseIf`.**
`If t < 0 Then t = 0 ElseIf t > 1 Then t = 1` is a compile "Syntax error";
split it into two single-line `If`s or a multi-line `If ... ElseIf ... End
If` block. Bit `PointSegDistM` (PR #24). Like the mid-module-declaration and
reserved-word traps, this only surfaces when the *module* is compiled, which
the build didn't force until `build\compile-check.ps1` was added (§9.1).

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

1. **NFC layer links** (cols 13 `COL_NFCAGOL` + 14 `COL_NFCMAP` since the
   PR #37 reorder) — state-aware (§9.3b), with descriptive labels
   ("Review NFC AGOL Layer" / "Review State NFC Layer"; Wisconsin splits
   into local-roads / state-trunk layer links, §7f). MI's primary link
   opens the FEMA-hosted NFC/ACUB webmap
   (`webmap=6a1702b9147243d1a5ee62cd614bc681` on `fema.maps.arcgis.com`)
   centered + markered on the row's coords — the MDOT Experience app's
   popup chrome blocked click-through, so the FEMA webmap replaced it.

2. **AGOL Map** (col 15 `COL_AGOLMAP` since PR #37) — driven by the
   `JobAgolMap` cell. When the user pastes their own AGOL webmap URL,
   every Sites row deep-links into it centered + markered on the row;
   when blank it defaults to the FEMA Map Viewer pin labeled
   "FEMA AGOL Map Viewer" (§7f).

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
- **IN/WI/MN/IL/OH** → `URL_NFC_MAPVIEW_IN/_WI/_MN/_IL/_OH` - side-load
  the state's own live NFC layer (§4.2a-§4.2e) via the Map
  Viewer's `url=` parameter, using the same `find=`/`marker=`/`level=`
  parameters already proven to work for `URL_FEMAVIEW` (rather than the
  `webmap=`+`center=` combination MI's link uses, which broke once
  before when a `visibleLayers` param was added - see the `URL_NFC_MAPVIEW`
  comment in `modConstants.bas`). No curated combined NFC+ACUB webmap
  exists for these states to point at instead. MN/IL/OH are MapServer
  (not FeatureServer) layers, which `url=` also accepts.
- **Anything else** → `URL_FEMAVIEW`, the plain pin+zoom with no data
  layer. A BLANK State cell shows the "Set State →" prompt in the cells
  (and `ClassifyRows` refuses to run, PR #36); the CSV export's resolver
  still falls back to MI links for that edge case only.

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

### 9.7 Service-URL overrides — the "lego swapout" (PR "WI layer swap")

Every REST endpoint the classifier queries can be re-pointed at runtime,
no code change and no rebuild, so a state moving or locking a layer (as
WisDOT just did to its local-roads service) is a paste-in-a-URL fix
instead of a code edit + rebuild + redeploy. The **same short KEY names**
are used in all four products so a swap is the same mental model
everywhere:

| Key | Default constant / REST entry |
|---|---|
| `MI_NFC` / `MI_ROUTE` | MDOT 353 / 543 |
| `IN_NFC` / `IN_ROADNAME` | INDOT LRSE 22 / centerlines 15 |
| `WI_LOCAL_ROADS` / `WI_STATE_TRUNK` | WisDOT local (queried first) / state-trunk |
| `ACUB` | NTAD Adjusted Urban Areas |
| `TIGER_ROADS` | Census TIGERweb local roads |
| `MN_NFC` / `IL_NFC` / `OH_NFC` | reference-only (unwired) |

- **Excel:** `modConstants.ServiceUrl(key)` returns the value pasted into
  the Sources sheet's "3. SERVICE URLs" table (workbook named range
  `Svc_<KEY>`, built by `modSources.SvcOverrideRow`) or the built-in
  `ServiceDefault(key)` when the cell is blank. Every `REST_*` usage in
  `modClassify` now goes through `ServiceUrl()`. Keep `ServiceDefault`,
  `SvcOverrideRow`'s row list, and the web/notebook default tables in
  sync.
- **Web:** rr-core defines `svc(key)` = `globalThis.RR_SERVICE_OVERRIDES[key]
  || REST[key]`; every `REST.X` read became `svc("X")`. A shared,
  byte-identical `<script>` block on **index.html** and **sources.html**
  (`#services`) loads pasted URLs from `localStorage["rr_service_overrides"]`
  into `globalThis.RR_SERVICE_OVERRIDES` and renders the editable table.
  rr-core stays self-contained/headless-safe (the override global is
  optional; `verify-web-core.mjs` never sets it, so it exercises
  defaults). NOTE: the two REST keys were renamed `MDOT_NFC`→`MI_NFC`,
  `MDOT_ROUTE`→`MI_ROUTE` to match the Excel/notebook key vocabulary.
- **Notebook:** `SERVICE_OVERRIDES` dict + `svc(key)`; same key names.
- **Scope:** overrides cover the classification *query* endpoints only,
  not the Map-Viewer deep-link templates (`URL_NFC_MAPVIEW*`) or the
  FIRMette GP service.

### 9.8 Map-page PDF export — Excel printable-area geometry (2026-07-14)

> **2026-07-16 — this whole pipeline is now the FALLBACK.** The Location Map
> PDF is written directly by `modPdf.bas` (§7g) because Excel's print
> RENDERER (as distinct from its paginator, which everything below concerns)
> was measured drawing the whole page vertically stretched ~7% on a real
> machine — no geometry below can fix that. Everything below stays true and
> load-bearing for the fallback path (WIA missing / image source file gone).

The "40 pages for 10 sites" / "images in slivers" / "giant whitespace" family
of bugs all trace to ONE fact, plus a stack of Excel quirks discovered while
fixing it. Recorded here because it is completely non-obvious and we may
revisit (e.g. if a FEMA laptop lacks Microsoft Print to PDF).

**The fact:** Excel paginates against the ACTIVE PRINTER's *usable area* =
paper − max(page-setup margin, driver hard margin) per side — even for
`ExportAsFixedFormat` PDF export, where no paper exists. Drivers reserve far
more than you'd guess, and it varies per machine:

| Driver (landscape Letter 792×612, margins 0) | measured usable |
|---|---|
| Brother HL-L2420DW (WSD) | 749 × 552 pt |
| **Microsoft Print to PDF** (Windows inbox) | **769.5 × 576 pt** |

Content ≥ usable is auto-split onto overflow pages (a 792×612 block became a
2×2 four-page tile). **Equality also loses** — device-unit rounding tips a
block sized exactly to the usable area over the edge, so strict slack is
mandatory. And **fit-to-page scaling is NOT a fix**: it silently discards the
manual `HPageBreaks` and reflows all blocks as one continuous ~80% strip
floating in whitespace (the symptom the user caught in a real export).

**The fix (modMaps.ExportCombinedMapPdf + modConstants):**

1. Blocks are sized to Microsoft Print to PDF's floor — `MAP_PAGE_WIDTH_PTS
   = 760`, `MAP_PAGE_HEIGHT_PTS = 568` (4 rows × 142pt) — strictly under
   769.5×576. Margins 0 + `CenterHorizontally/Vertically` leave a 16pt side /
   22pt top-bottom frame. (True borderless is impossible: every driver,
   including MS PDF, reserves an edge.)
2. The export **temporarily switches `Application.ActivePrinter` to
   "Microsoft Print to PDF"** so the pagination uses that machine-constant
   inbox driver, then restores the user's printer (also on the Fail path).
   Port resolution: read `HKCU\Software\Microsoft\Windows NT\CurrentVersion\
   Devices\Microsoft Print to PDF` → `"winspool,Ne0X:"` → `"Microsoft Print
   to PDF on Ne0X:"`; probing Ne00..Ne31 is only the fallback (port numbers
   shuffle between sessions — a hardcoded Ne00 worked once by luck).
3. Excel traps hit on the way (each cost a debugging round):
   - `ActivePrinter` **cannot be set when no workbook is open** (a probe
     doing so silently kept the old printer and mismeasured).
   - **Switching printers resets the sheet's page setup** — orientation
     flipped back to portrait. `ConfigureMapPageSetup` + `SetMapPrintArea`
     must re-run AFTER the switch, before export.
   - `modMapImage.PageTopPts` must compute page tops with the same
     `MAP_FIRST_PAGE_ROW` offset as `modMaps.CreateMapPage` — when the
     header band moved pages down, the stale formula piled every inserted
     image at the top of the sheet, overlapping (each printed page showed
     slivers of several screenshots).
   - `IMG_INSET_PTS = 1`: a picture ending exactly ON a page-break line gets
     tipped onto an overflow page by rounding (12 pages for 6 sites — full
     page + sliver page alternating).
4. Fallback when MS Print to PDF is absent (it's a removable Windows
   feature): `FitToPagesWide=1 / FitToPagesTall=MapPageCount` — right page
   count, shrunken content with whitespace; better than a spill.

**The measurement technique (re-run this if pages ever split again):** build a
probe sheet — 200 rows × 6pt with a value in each, 50 narrow columns, target
page setup, desired printer active — and read `ws.HPageBreaks(1).Location.Row`
/ `ws.VPageBreaks(1).Location.Column`. (rows−1)×6pt = the driver's true usable
height; likewise width. That is the ceiling any block must stay strictly under.

**Verification done:** end-to-end with the user's six real Google Earth
screenshots on a copy of the built inspector — `PrepareMapPages` →
`InsertMapImages` → `ExportCombinedMapPdf` produced exactly 6 PDF pages, one
near-full-bleed image per page, stamp on top, default printer restored after.
Page counts and placements checked with PyMuPDF (`fitz`); note
`page.get_image_rects()` returns the **uncropped** image extents, which
mislead — trust page count + rendered pixels, not those rects.

