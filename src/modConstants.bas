Attribute VB_Name = "modConstants"
Option Explicit

' RoadReviewer / Site Inspector Review Tool - shared constants.
' Single place for product ids, sheet names, Sites column indices,
' named-range names, URL templates and the FHWA FunctionalSystem domain.
' Friction point #9 (mixing column letters and indices) is resolved by
' using 1-based column index constants everywhere.
'
' Two products are built from this one src/ tree (see build\build.ps1):
'   Standard  -> RoadReviewer.xlsm            (PDMGs / partners / reviewers)
'   Inspector -> Site Inspector Review Tool.xlsm (full site-inspector kit)
' The product id is baked into the workbook at build time as a hidden
' defined name (NM_PRODUCT) via modUtil.SetProduct.

' ---- Products ----
Public Const PRODUCT_STANDARD As String = "Standard"
Public Const PRODUCT_INSPECTOR As String = "Inspector"
' Hidden workbook-level defined name that stores which product this file is.
Public Const NM_PRODUCT As String = "RR_Product"

' Version stamp shown on Start Here + the Sources sheet so a shared copy can
' be traced back to the PR / build it came from. Bump this on each release.
Public Const BUILD_REFERENCE As String = "PR #37"

' ---- Sheet names ----
' The "hub" sheet name. On the standard product this is the visible landing
' ("Start Here"); on the inspector it's the hidden utility sheet SH_TOOLS. Use
' modUtil.StartSheetName() to get the right one for the current product; the two
' constants below are just the literal names.
Public Const SH_START As String = "Start Here"
Public Const SH_TOOLS As String = "Tools and Exports"
Public Const SH_SITES As String = "Sites"
Public Const SH_SOURCES As String = "Sources"
' The map workspace. On the inspector this is the LANDING sheet; on the standard
' product it ships hidden (opt-in). Tab reads "Map Pages".
Public Const SH_MAPPAGES As String = "Map Pages"

' ---- Start Here named ranges ----
' WO/DI/Disaster/Applicant/Buffer are inspector-only; the standard product
' never creates them. Shared code reads them via modUtil.SetupValue, which
' returns "" for a missing name, so every consumer already degrades cleanly.
Public Const NR_WO As String = "JobWO"
Public Const NR_DI As String = "JobDI"
Public Const NR_DISASTER As String = "JobDisaster"
Public Const NR_APPLICANT As String = "JobApplicant"
Public Const NR_STATE As String = "JobState"
Public Const NR_OUTFOLDER As String = "JobOutputFolder"
' Optional - when the user pastes an ArcGIS Online webmap URL here, the
' Sites table's "AGOL Map" column (COL_AGOLMAP) produces per-row deep-links
' that open the webmap centered on the row's coordinates.
Public Const NR_AGOLMAP As String = "JobAgolMap"
' Optional imagery-source override for Fetch Imagery, entered right on the
' Map Pages header band. Blank = Esri World Imagery (REST_WORLD_IMAGERY /
' the Svc_WORLD_IMAGERY override). A pasted URL must be an ArcGIS MapServer
' - only MapServers expose the /export operation Fetch Imagery drives;
' Query-only FeatureServers cannot render an image (same constraint as the
' web tool's PDF figures, §7b).
Public Const NR_IMAGERYSVC As String = "JobImagerySvc"
' Search radius (in feet) used when the exact point-on-polyline intersect
' returns no road segments. The classifier always tries an exact intersect
' first; this buffer is the second-chance fallback. Default 250 ft, read
' from an editable Start Here cell on BOTH products, capped to [1, 1000] in
' modClassify.BufferFeet.
Public Const NR_BUFFER As String = "JobBufferFeet"
Public Const DEFAULT_BUFFER_FEET As Long = 250

' ---- Sites table geometry ----
' Row 1 IS the header row. The action buttons that used to float over a slim
' row-1 toolbar have moved back to Start Here (they were free-floating shapes
' anchored at Top:=2pt, so they would have covered the headers here).
Public Const SITES_HEADER_ROW As Long = 1
Public Const SITES_FIRST_DATA_ROW As Long = 2
Public Const SITES_FORMULA_ROWS As Long = 500   ' hyperlink/validation pre-fill depth

' ---- Sites column indices (1-based) ----
' Order is optimized for copy-paste: Latitude | Longitude | Description are
' contiguous so a three-column paste from any spreadsheet lands in one go.
' Address sits to the RIGHT of Description because it is rarely used (it
' only feeds the geocoder when lat/lon are blank).
Public Const COL_WO As Long = 1          ' inspector-only (hidden in standard)
Public Const COL_DI As Long = 2          ' inspector-only (hidden in standard)
Public Const COL_SITENO As Long = 3
Public Const COL_SITENAME As Long = 4
Public Const COL_LAT As Long = 5
Public Const COL_LON As Long = 6
Public Const COL_DESC As Long = 7
Public Const COL_ADDRESS As Long = 8
Public Const COL_CATEGORY As Long = 9
' Two optional user-data columns. Both flow into the KML description tag,
' the MapPages textbox stamp, and the CSV export - anywhere Description
' shows up - when they have values.
Public Const COL_COSTS As Long = 10
Public Const COL_WORKCOMP As Long = 11
Public Const COL_GEOCODE As Long = 12
' ---- map links (13-15, reordered PR #37) ----
' PRIMARY map link. Opens the state functional-class layer in ArcGIS Map
' Viewer, CENTERED + MARKERED on the row's point. This is the link that
' actually navigates to the coordinate, so it leads. For Wisconsin this is
' the LOCAL ROADS layer ("Review Local Roads Layer"); every other state gets
' its NFC layer ("Review NFC AGOL Layer").
Public Const COL_NFCAGOL As Long = 13
' The state's second reference link, right next to the first (user request:
' the two map-layer links sit together as M and N). For Wisconsin this is
' the STATE TRUNK highway layer in Map Viewer ("Review State Trunk Hwy
' Layer"); every other state opens its official public app at its default
' extent ("Review State NFC Layer" - the Experience apps mis-navigate on
' coordinate deep-links, PR #17/#18).
Public Const COL_NFCMAP As Long = 14
' The AGOL webmap link. When the user pastes their own webmap URL
' (NR_AGOLMAP) every row deep-links into it centered on the point; when the
' field is BLANK it defaults to the FEMA-hosted ArcGIS Map Viewer pin
' ("FEMA AGOL Map Viewer", PR #37) - which is why the separate FEMA Viewer
' column now ships hidden.
Public Const COL_AGOLMAP As Long = 15

' ---- classification results (contiguous 16-22, reordered PR #37) ----
' Written only by CheckRoads / ReRunFailedRows. HIDDEN until one of those
' STARTS (revealed before the loop so the user watches them fill in). The
' verdict columns lead the block per user request: Federal Aid Status and
' Review Reason are the first two (P and Q on the sheet), the supporting
' lookups follow. The grey tint + tri-color verdict conditional format both
' span this range.
Public Const COL_ELIGIBILITY As Long = 16
' <=3-word reason a row is yellow ("Review"): Nearby FHWA road / Second road
' close / Urban boundary edge / etc. Blank for confident red or green rows.
Public Const COL_REVIEWNOTE As Long = 17
Public Const COL_CLASS As Long = 18
Public Const COL_URBANRURAL As Long = 19
Public Const COL_ACUBNAME As Long = 20
Public Const COL_ROADNAME As Long = 21
' Street name(s) from the U.S. Census Bureau's TIGERweb Local Roads layer.
' Dense (covers most addresses) where COL_ROADNAME / MDOT 543 only gets
' trunkline route designations like "I-94 BL". Pipe-joined when multiple
' streets fall inside the search buffer (intersection points).
Public Const COL_STREET As Long = 22

' First and last of the auto-reviewer block, so the hide/show helper and the
' conditional formats never drift out of sync with the constants above.
Public Const COL_REVIEWER_FIRST As Long = COL_ELIGIBILITY
Public Const COL_REVIEWER_LAST As Long = COL_STREET

' ---- imagery / photo-source links (23-28) ----
Public Const COL_GMAP As Long = 23
Public Const COL_STREETVIEW As Long = 24
Public Const COL_BING As Long = 25
Public Const COL_GEARTH As Long = 26
Public Const COL_FEMAVIEW As Long = 27
Public Const COL_FIRMPORTAL As Long = 28
Public Const COL_FIRMSTATUS As Long = 29  ' inspector-only (hidden in standard)
Public Const COL_MAPSTATUS As Long = 30   ' inspector-only (hidden in standard)
Public Const COL_LAST As Long = 30

' ---- Verification map URL templates (§4.3). {LAT}/{LON} substituted at run time. ----
Public Const URL_GMAP As String = "https://www.google.com/maps?q={LAT},{LON}"
Public Const URL_STREETVIEW As String = "https://www.google.com/maps?q&layer=c&cbll={LAT},{LON}"
Public Const URL_BING As String = "https://www.bing.com/maps?cp={LAT}~{LON}&lvl=18&style=h"
' Google Earth web - includes the historical-imagery slider, the main
' pre-disaster-condition source users asked for by name.
Public Const URL_GEARTH As String = "https://earth.google.com/web/search/{LAT},{LON}"
Public Const URL_FEMAVIEW As String = "https://fema.maps.arcgis.com/apps/mapviewer/index.html?find={LON}%2C{LAT}&marker={LON},{LAT},4326&level=16"
Public Const URL_FIRMPORTAL As String = "https://msc.fema.gov/portal/firmette?latitude={LAT}&longitude={LON}"
' Open the row's point on the FEMA-hosted ArcGIS Map Viewer with MDOT's NFC
' FeatureServer (layer 353) side-loaded and a clean pin (find/marker/level) -
' the same pattern as the Indiana/Wisconsin links below.
' History: this used to open MDOT's curated NFC/NHS/ACUB webmap
' (item 6a1702b9...), but that webmap carries a "Metropolitan Planning
' Organizations" layer whose region labels read as the site name.
'
' Michigan's "AGOL NFC Layer" column (COL_NFCAGOL) uses THIS webmap
' (item 6a1702b9...), centered + markered on the point. It shows functional
' class + adjusted urban boundary in ArcGIS Map Viewer and - unlike a raw
' side-load of the time-enabled layer 353 - does NOT surface a time slider
' (the webmap is preconfigured), which fixes the "timeline filter hides the
' roads" problem (PR #25). The MPO-region label can still appear here, but
' that's tolerable on this secondary column now that the main "Open" column
' (COL_NFCMAP) goes to the official public app instead. [NEEDS a browser
' click-test - ArcGIS Map Viewer time behavior can't be verified headless.]
Public Const URL_NFC_MAPVIEW As String = "https://fema.maps.arcgis.com/apps/mapviewer/index.html?webmap=6a1702b9147243d1a5ee62cd614bc681&center={LON},{LAT}&level=16&marker={LON},{LAT}"

' Indiana and Wisconsin don't have a curated combined NFC+ACUB webmap the
' way MI does, so these side-load the state's own live NFC FeatureServer
' (§4.2a/§4.2b) directly into the same FEMA Map Viewer base URL + params
' already proven working for URL_FEMAVIEW above (find=/marker=/level=)
' rather than the webmap=+center= combination that broke when a
' visibleLayers param was added (see the note above). Degrades gracefully
' even if the Map Viewer version in use doesn't render the `url=` layer -
' the pin+zoom still resolves correctly either way.
Public Const URL_NFC_MAPVIEW_IN As String = "https://fema.maps.arcgis.com/apps/mapviewer/index.html?url=https%3A%2F%2Fgisdata.in.gov%2Fserver%2Frest%2Fservices%2FHosted%2FLRSE_Functional_Class%2FFeatureServer%2F22&find={LON}%2C{LAT}&marker={LON},{LAT},4326&level=16"
Public Const URL_NFC_MAPVIEW_WI As String = "https://fema.maps.arcgis.com/apps/mapviewer/index.html?url=https%3A%2F%2Fservices5.arcgis.com%2F0pgGLzT0Nh7FVjon%2Farcgis%2Frest%2Fservices%2FFFCL_gdb%2FFeatureServer%2F3&find={LON}%2C{LAT}&marker={LON},{LAT},4326&level=16"
' Wisconsin's LOCAL ROADS layer side-loaded the same way (PR #37): WI's
' primary link column reviews the local-roads layer (it carries most points),
' with the state-trunk link above demoted to the second column.
' NB: the url= value is deliberately NOT percent-encoded here - encoded, the
' full link is ~261 chars and Excel's HYPERLINK() errors (#VALUE!) past 255.
' Unencoded is legal (the value contains no "&") and keeps the link at ~239.
Public Const URL_NFC_MAPVIEW_WI_LOCAL As String = "https://fema.maps.arcgis.com/apps/mapviewer/index.html?url=https://services5.arcgis.com/0pgGLzT0Nh7FVjon/arcgis/rest/services/Functional_Class_Local_Non_Prod/FeatureServer/1&find={LON}%2C{LAT}&marker={LON},{LAT},4326&level=16"
' MN/IL/OH (PR #36): same side-load shape. These are MapServer layers rather
' than FeatureServers, which Map Viewer's url= parameter also accepts.
Public Const URL_NFC_MAPVIEW_MN As String = "https://fema.maps.arcgis.com/apps/mapviewer/index.html?url=https%3A%2F%2Fdotapp9.dot.state.mn.us%2Fegis12%2Frest%2Fservices%2FBASEMAP%2Fmndot_commonlayers2%2FMapServer%2F11&find={LON}%2C{LAT}&marker={LON},{LAT},4326&level=16"
Public Const URL_NFC_MAPVIEW_IL As String = "https://fema.maps.arcgis.com/apps/mapviewer/index.html?url=https%3A%2F%2Fgis1.dot.illinois.gov%2Farcgis%2Frest%2Fservices%2FAdministrativeData%2FFunctionalClass%2FMapServer%2F0&find={LON}%2C{LAT}&marker={LON},{LAT},4326&level=16"
Public Const URL_NFC_MAPVIEW_OH As String = "https://fema.maps.arcgis.com/apps/mapviewer/index.html?url=https%3A%2F%2Ftims.dot.state.oh.us%2Fags%2Frest%2Fservices%2FRoadway_Information%2FFunctional_Class%2FMapServer%2F0&find={LON}%2C{LAT}&marker={LON},{LAT},4326&level=16"

' Official public-facing functional-class apps per state (verified live
' 2026-07-05, §7c). These drive the "NFC Map" / Open column (COL_NFCMAP) -
' the authoritative public app, distinct from the ArcGIS Map Viewer layer
' links above. They are plain app URLs with NO {LAT}/{LON}: the state
' Experience apps can't be reliably centered on a point via URL (the
' coordinate deep-link mis-navigates, PR #17/#18), so a row's Open link
' opens the app at its default extent and the AGOL NFC Layer column is what
' centers on the exact point. WI has no statewide interactive app, so its
' Open link falls back to WisDOT's official functional-class page.
Public Const APP_MI As String = "https://experience.arcgis.com/experience/7edd160c205d46b481fcd605bb4c58ce"
Public Const APP_IN As String = "https://experience.arcgis.com/experience/e388c2aa14aa4788a702705620567589/?org=indot"
Public Const APP_WI As String = "https://wisconsindot.gov/Pages/projects/data-plan/plan-res/function.aspx"
Public Const APP_MN As String = "https://webgis.dot.state.mn.us/emma/"
Public Const APP_IL As String = "https://www.gettingaroundillinois.com/MapViewer/?config=RFCconfig.json"
Public Const APP_OH As String = "https://tims.dot.state.oh.us/tims"

' ---- REST endpoints (§4.1, §4.2, §8.2) ----
Public Const REST_FIRMETTE As String = "https://msc.fema.gov/arcgis/rest/services/NFHL_Print/MSCPrintB/GPServer/PrintFIRMette"
Public Const REST_MDOT_NFC As String = "https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/353"
Public Const REST_MDOT_ROUTE As String = "https://mdotgis.state.mi.us/arcgis/rest/services/Widget/NextGenPrFinderPub/FeatureServer/543"
Public Const REST_ACUB As String = "https://services.arcgis.com/xOi1kZaI0eWDREZv/arcgis/rest/services/NTAD_Adjusted_Urban_Areas/FeatureServer/0"

' Indiana NFC (§4.2a). LRSE_Functional_Class carries the bare FHWA 1-7 code
' (field functional_class) with no urban/rural embedded - same shape as MDOT
' 353. Road name is NOT on this layer (no shared LRS join key), so it's a
' separate point-intersect query against the statewide centerline layer.
Public Const REST_IN_NFC As String = "https://gisdata.in.gov/server/rest/services/Hosted/LRSE_Functional_Class/FeatureServer/22"
Public Const REST_IN_ROADNAME As String = "https://gisdata.in.gov/server/rest/services/Hosted/Road_Centerlines_of_Indiana_2021/FeatureServer/15"

' Wisconsin NFC (§4.2b) - two layers. The LOCAL Road Network layer
' (county/city/town roads AND most collectors) is queried FIRST because it
' covers the large majority of points; the State Trunk Network
' (interstates/state highways) is queried only when the local layer has no
' classification for the point (§4.2b query strategy, PR "WI layer swap").
'   - State Trunk: bare FHWA 1-7 code in FED_FC_CD (string), plus its own
'     URB_TYPE field (unused - ACUB is the single urban/rural source per §4.2).
'   - Local: FNCT_CLS_CTGY_TYCD encodes urban/rural INTO the class code and
'     needs WisconsinLocalCategoryToFhwa() below to normalize to bare FHWA 1-7.
'     State highways also appear in this layer but with a NULL/0 class (an
'     unclassified "stub"); those are skipped and trigger the trunk fallback.
' Link note (PR "WI layer swap"): WisDOT locked/renamed the old local-roads
' snapshot (WI_Local_Roads_Flood_Damage_Assessment_Snapshot - now token-gated);
' the live public local layer is Functional_Class_Local_Non_Prod/1. If WisDOT
' moves it again, paste the new URL into the Sources sheet's Service URLs table
' (named range Svc_WI_LOCAL_ROADS) - no rebuild needed (ServiceUrl below).
Public Const REST_WI_STATE_TRUNK As String = "https://services5.arcgis.com/0pgGLzT0Nh7FVjon/arcgis/rest/services/FFCL_gdb/FeatureServer/3"
Public Const REST_WI_LOCAL_ROADS As String = "https://services5.arcgis.com/0pgGLzT0Nh7FVjon/arcgis/rest/services/Functional_Class_Local_Non_Prod/FeatureServer/1"
Public Const REST_CENSUS_GEOCODE As String = "https://geocoding.geo.census.gov/geocoder/locations/onelineaddress"
' U.S. Census Bureau TIGERweb — Local Roads (full detail layer 8 of the
' Transportation MapServer). Returns the NAME field for any street whose
' centerline intersects the search buffer. Covers every state, free, no
' auth. We hit this in addition to MDOT 543 because MDOT only carries
' designated trunkline routes — TIGER fills in the local street names.
Public Const REST_TIGER_ROADS As String = "https://tigerweb.geo.census.gov/arcgis/rest/services/TIGERweb/Transportation/MapServer/8"

' Minnesota / Illinois / Ohio NFC class services - WIRED in PR #36 (§4.2c/d/e).
' All three publish a bare FHWA 1-7 class code, same shape as Indiana, with no
' retired-segment filter needed (confirmed live 2026-07-15 via ?f=pjson +
' returnDistinctValues):
'   MN: FUNCTIONAL_CLASS (integer 1-7; FUNCTIONAL_CLASS_DESC carries the
'       matching FHWA label). No road-name field (ROUTE_ID is an LRS key) -
'       names come from the Census TIGER backfill.
'   IL: FC (STRING "1".."7", published coded-value domain). Route-system
'       labels like "FAU 1422" exist but are not street names - TIGER covers.
'   OH: FUNCTION_CLASS_CD (integer 1-7). ROUTE_TYPE+ROUTE_NBR give trunkline
'       names ("US 23", "SR 161") for IR/US/SR routes; municipal "MR" codes
'       are skipped (cryptic) and TIGER fills local names.
Public Const REST_MN_NFC As String = "https://dotapp9.dot.state.mn.us/egis12/rest/services/BASEMAP/mndot_commonlayers2/MapServer/11"
Public Const REST_IL_NFC As String = "https://gis1.dot.illinois.gov/arcgis/rest/services/AdministrativeData/FunctionalClass/MapServer/0"
Public Const REST_OH_NFC As String = "https://tims.dot.state.oh.us/ags/rest/services/Roadway_Information/Functional_Class/MapServer/0"

' Service-URL overrides ("lego swapout", PR "WI layer swap"): every REST
' endpoint the classifier queries can be re-pointed at runtime via the Sources
' sheet's "Service URLs" table. The resolver functions ServiceUrl/ServiceDefault
' live at the BOTTOM of this module (VBA requires all module-level Const/Dim
' declarations to precede every procedure, §9.3 - so the functions can't sit
' here between the Const blocks).

' MDOT requires a browser User-Agent or it returns HTTP 403 (§4.2 operational note).
Public Const BROWSER_UA As String = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"

' ---- State selector (F8). All six Region V states are wired for road
' classification (MN/IL/OH joined in PR #36). The ACUB-only path still exists
' in modClassify for any OTHER typed state code, and modUtil.BareStateCode
' still strips a legacy "(not wired)" suffix so an older workbook whose cell
' holds "MN (not wired)" keeps working. ----
Public Const STATE_LIST As String = "WI,IN,MI,MN,IL,OH"

' ---- Status-prefix used by the "re-run failed rows" feature (F12). ----
Public Const STATUS_FAILED_PREFIX As String = "Failed - "

' ---- FEMA Print FIRMette GP-job polling (inspector product) ----
Public Const FIRMETTE_POLL_INTERVAL_SECONDS As Long = 2
Public Const FIRMETTE_POLL_MAX_ATTEMPTS As Long = 90    ' ~3 min max

' ---- MapPages layout (inspector product, ported from prototype) ----
' Each map page = 4 rows x 13 cols, sized to fit INSIDE a landscape-Letter page
' (792x612) with a fixed MAP_PRINT_MARGIN_PTS frame on all sides. Excel bases
' its printable area on the default printer's hard margins even for PDF export
' and ALWAYS reserves an epsilon - true edge-to-edge content overflows on some
' drivers (observed: a 612-tall block spilling into a 2x2 four-page block on a
' Brother laser, and a 24-page spill even on Microsoft Print to PDF). A 0.25"
' margin clears every real driver's hard margins, so the export is 1:1
' (Zoom=100), the manual page breaks hold, and each map page is EXACTLY one
' PDF page with a thin even frame.
' Sizing is driven by the MEASURED usable print area of "Microsoft Print to
' PDF", which ExportCombinedMapPdf makes the active printer for the export (and
' restores after). Excel's usable area = paper - max(page-setup margin, DRIVER
' hard margin) per side; a probe sheet (thin rows/cols, find the first
' automatic page break) measured MS Print to PDF's landscape-Letter usable at
' 769.5 x 576pt - no Windows driver is truly borderless, but this one is the
' same on EVERY machine (Windows inbox driver), which makes the geometry
' machine-independent. Any block sized AT or past the usable floor is
' auto-split into overflow pages (the 40-pages-for-10-sites / strip-page
' bugs; equality loses to device rounding, so slack is mandatory).
' 760 x 568 sits under the floor; CenterHorizontally/Vertically turn the
' leftover into an even frame (16pt sides, 22pt top/bottom).
Public Const MAP_ROWS_PER_PAGE As Long = 4
Public Const MAP_COLS_WIDE As Long = 13
Public Const MAP_PRINT_MARGIN_PTS As Double = 0             ' page-setup margin
Public Const MAP_PAGE_HEIGHT_PTS As Double = 568            ' 4 rows x 142pt, < 576 usable
Public Const MAP_PAGE_WIDTH_PTS As Double = 760             ' < 769.5 usable
Public Const MAP_TEXTBOX_WIDTH As Double = 260
Public Const MAP_TEXTBOX_HEIGHT As Double = 104
' Stamp font size (WO/DI/Applicant/Site/coords textbox on each page). Bumped
' from 9 for legibility on the printed page.
Public Const MAP_STAMP_FONT As Double = 11

' Per-page "Select photo" picker button (modMapImage.PickImageForPage). One is
' stamped on each map page by CreateMapPage; hidden during PDF export.
Public Const MAP_PICKBTN_PREFIX As String = "PickBtn_Page_"

' ---- Fetch Imagery (modMapFetch) ----
' Esri World Imagery export endpoint - one anonymous GET returns a rendered
' aerial PNG for a Web-Mercator bbox (confirmed live 2026-07-14: HTTP 200,
' image/png, real 1520x1136 satellite image). No auth, no browser needed.
' Reachability from a hardened FEMA laptop is UNTESTED; if it's blocked there,
' the Sources sheet's Service URLs table (key WORLD_IMAGERY) can re-point it.
Public Const REST_WORLD_IMAGERY As String = "https://services.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer"
' Half-WIDTH of the fetched frame in meters (bbox = site center +/- this).
' The frame height is scaled down to the page-block aspect ratio
' (MAP_PAGE_WIDTH_PTS : MAP_PAGE_HEIGHT_PTS = 760 : 568), so the site is
' always the exact center of the printed image - which is what lets
' modMapFetch drop the site pin at the geometric center of the page.
Public Const MAP_IMG_HALFWIDTH_M As Double = 600
' Requested raster size: 2x the page block's point size (760x568 -> 1520x1136)
' so the placed image stays sharp at print resolution.
Public Const MAP_IMG_PX_W As Long = 1520
Public Const MAP_IMG_PX_H As Long = 1136
' Esri's terms require visible attribution on maps produced from this service.
' modMapFetch prints this line in a small textbox on every fetched page.
Public Const MAP_IMG_ATTRIBUTION As String = "Imagery: Esri, Maxar, Earthstar Geographics"
' Shape-name prefixes for the printed site pin + attribution line. Deliberately
' NOT MAP_CTRL_PREFIX / MAP_PICKBTN_PREFIX: SetMapEditControlsVisible hides
' those groups during PDF export, and the pin + attribution MUST print.
Public Const MAP_PIN_PREFIX As String = "MapPin_Page_"
Public Const MAP_ATTR_PREFIX As String = "MapAttr_Page_"

' ---- MapPages job panel ----
' MapPages is now a PERMANENT sheet built by BuildWorkbook (it used to be
' created - and deleted - by PrepareMapPages). It hosts the job inputs so the
' inspector fills WO/DI/Disaster/Applicant/Output Folder on the same sheet the
' pages live on, instead of bouncing back to Start Here.
'
' The page rows are ~153pt tall (4 of them = one landscape page), and a row
' height is a whole-row property - so input cells sitting beside the pages would
' be 153pt tall too. Hence a band of SHORT rows above the pages: rows
' 1..MAP_HEADER_ROWS hold the header (workflow ribbon + job info + FIRMettes),
' and page 1 starts at MAP_FIRST_PAGE_ROW. The pages' print area starts there
' too, so the header never prints. Buttons are absolute-positioned shapes over
' this band, so they're free of the tall page-row heights.
Public Const MAP_HEADER_ROWS As Long = 22
Public Const MAP_FIRST_PAGE_ROW As Long = MAP_HEADER_ROWS + 1
Public Const MAP_HEADER_ROW_HEIGHT As Double = 17
' Job-block cells inside that band. Labels merge A:B, values merge C:F (the page
' columns are only ~61pt wide each, so a single cell is too narrow for a label).
Public Const MAP_JOB_LABEL_COL As Long = 1      ' A (merged A:B)
Public Const MAP_JOB_VALUE_COL As Long = 3      ' C (merged C:F)
Public Const MAP_JOB_VALUE_LAST_COL As Long = 6 ' F
' The job block sits under the one-click hero button + the collapsed Advanced
' section (which together occupy roughly rows 1-12 of the band; the "Job info"
' label sits on row 13).
Public Const MAP_JOB_FIRST_ROW As Long = 14     ' WO # ... Imagery URL run 14..20

' Returns the FHWA functional-class label for a bare FHWA 1-7 code. MDOT
' (layer 353, LrseFunctionalSystem domain), Indiana (LRSE_Functional_Class,
' dFunctionalClass domain) and Wisconsin (FED_FC_CD, plus the local-roads
' layer once normalized by WisconsinLocalCategoryToFhwa) all share this same
' numeric FHWA standard, so one label table covers all three wired states
' (§4.2). Code 0 only comes out of Michigan's domain ("Non-Certified
' Roadway") or an unrecognized Wisconsin local-road category - both fall
' into FederalAidVerdict's "review manually" bucket rather than a verdict.
Public Function FunctionalSystemLabel(ByVal code As Long) As String
    Select Case code
        Case 0: FunctionalSystemLabel = "Non-Certified Roadway"
        Case 1: FunctionalSystemLabel = "Interstate"
        Case 2: FunctionalSystemLabel = "Other Freeway"
        Case 3: FunctionalSystemLabel = "Other Principal Arterial"
        Case 4: FunctionalSystemLabel = "Minor Arterial"
        Case 5: FunctionalSystemLabel = "Major Collector"
        Case 6: FunctionalSystemLabel = "Minor Collector"
        Case 7: FunctionalSystemLabel = "Local"
        Case Else: FunctionalSystemLabel = "Unknown (" & code & ")"
    End Select
End Function

' Wisconsin's local-roads layer (REST_WI_LOCAL_ROADS) encodes urban/rural
' directly into FNCT_CLS_CTGY_TYCD instead of carrying a bare FHWA code
' (confirmed live via the layer's renderer uniqueValueInfos, §4.2b):
'   10 Rural Principal Arterial   60 Urban Principal Arterial
'   20 Rural Minor Arterial       86 Urban Minor Arterial Other
'   30 Rural Major Collector      96 Urban Collector Other
'   40 Rural Minor Collector
'   45 Rural Local                97 Urban Local
' "Urban Collector Other" (96) doesn't distinguish major vs minor collector.
' That split only changes the federal-aid verdict for RURAL collectors (rural
' major = federal aid, rural minor = not) - every *urban* collector is federal
' aid regardless of major/minor, so mapping 96 to Major Collector (5) is safe
' either way. ACUB (not this field) remains the source of truth for the
' Urban/Rural column; the rural/urban half of each code is discarded here.
Public Function WisconsinLocalCategoryToFhwa(ByVal categoryCode As Long) As Long
    Select Case categoryCode
        Case 10, 60: WisconsinLocalCategoryToFhwa = 3    ' Principal Arterial
        Case 20, 86: WisconsinLocalCategoryToFhwa = 4    ' Minor Arterial
        Case 30, 96: WisconsinLocalCategoryToFhwa = 5    ' Major Collector (96 = Urban Collector, major/minor unresolved)
        Case 40: WisconsinLocalCategoryToFhwa = 6         ' Rural Minor Collector
        Case 45, 97: WisconsinLocalCategoryToFhwa = 7     ' Local
        Case Else: WisconsinLocalCategoryToFhwa = 0       ' unrecognized - flagged for manual review
    End Select
End Function

' ---- Service-URL overrides ("lego swapout", PR "WI layer swap") ----
' Every REST endpoint the classifier queries can be re-pointed at runtime by
' pasting a replacement service URL into the matching cell on the Sources sheet
' (named range "Svc_<KEY>", built by modSources). This lets a user recover
' instantly when a publisher moves or locks a layer - as WisDOT just did to its
' local-roads snapshot - with no code change and no rebuild. The SAME key names
' are used by the web tool (RR_SERVICE_OVERRIDES) and the AGOL notebook
' (SERVICE_OVERRIDES) so one mental model spans every product. ServiceUrl(key)
' returns the pasted override when present, else the ServiceDefault(key).
Public Function ServiceUrl(ByVal key As String) As String
    Dim pasted As String
    pasted = SetupValue("Svc_" & key)
    If Len(pasted) > 0 Then
        ServiceUrl = pasted
    Else
        ServiceUrl = ServiceDefault(key)
    End If
End Function

' The built-in default endpoint for each override key. Keep this list in sync
' with modSources.SvcOverrideRow (the Sources-sheet table) and with the web /
' notebook default tables.
Public Function ServiceDefault(ByVal key As String) As String
    Select Case key
        Case "MI_NFC":         ServiceDefault = REST_MDOT_NFC
        Case "MI_ROUTE":       ServiceDefault = REST_MDOT_ROUTE
        Case "IN_NFC":         ServiceDefault = REST_IN_NFC
        Case "IN_ROADNAME":    ServiceDefault = REST_IN_ROADNAME
        Case "WI_STATE_TRUNK": ServiceDefault = REST_WI_STATE_TRUNK
        Case "WI_LOCAL_ROADS": ServiceDefault = REST_WI_LOCAL_ROADS
        Case "ACUB":           ServiceDefault = REST_ACUB
        Case "TIGER_ROADS":    ServiceDefault = REST_TIGER_ROADS
        Case "WORLD_IMAGERY":  ServiceDefault = REST_WORLD_IMAGERY
        Case "MN_NFC":         ServiceDefault = REST_MN_NFC
        Case "IL_NFC":         ServiceDefault = REST_IL_NFC
        Case "OH_NFC":         ServiceDefault = REST_OH_NFC
        Case Else:             ServiceDefault = ""
    End Select
End Function
