Attribute VB_Name = "modSources"
Option Explicit

' Sources & Quirks sheet. Two sections, per user direction (PR #22):
'   1. SOURCES   - for each state, the official public map (verify a point on
'                  it) and the data service this tool reads, as short
'                  bluebook-style citations ending "available at: <url>".
'   2. QUIRKS & CAVEATS - the same schema quirks in plain language.
' Official-app + unwired-state REST URLs were verified live 2026-07-05
' (see the research report in the PR). Wired/nationwide services reuse the
' REST_* constants the classifier actually queries. Keep this sheet in sync
' with CLAUDE.md §4.2/§4.2a/§4.2b and web/sources.html.

Private mRow As Long   ' running row cursor while writing the sheet

' Official public-facing functional-class maps (verified live 2026-07-05).
Private Const APP_MI As String = "https://experience.arcgis.com/experience/7edd160c205d46b481fcd605bb4c58ce"
Private Const APP_IN As String = "https://experience.arcgis.com/experience/e388c2aa14aa4788a702705620567589/?org=indot"
Private Const APP_WI As String = "https://wisconsindot.gov/Pages/projects/data-plan/plan-res/function.aspx"
Private Const APP_MN As String = "https://webgis.dot.state.mn.us/emma/"
Private Const APP_IL As String = "https://www.gettingaroundillinois.com/MapViewer/?config=RFCconfig.json"
Private Const APP_OH As String = "https://tims.dot.state.oh.us/tims"

' Reference-only REST services for the not-yet-wired states (we don't query
' these; they're cited so a future release can wire them and so a reviewer
' knows where the authoritative data lives).
Private Const REST_MN_NFC As String = "https://dotapp9.dot.state.mn.us/egis12/rest/services/BASEMAP/mndot_commonlayers2/MapServer/11"
Private Const REST_IL_NFC As String = "https://gis1.dot.illinois.gov/arcgis/rest/services/AdministrativeData/FunctionalClass/MapServer/0"
Private Const REST_OH_NFC As String = "https://tims.dot.state.oh.us/ags/rest/services/Roadway_Information/Functional_Class/MapServer/0"

Public Sub BuildSourcesSheet()
    Dim ws As Worksheet
    Set ws = FreshSheet(SH_SOURCES)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 120
    mRow = 2

    With ws.Cells(mRow, 2)
        .Value = "Sources & Quirks"
        .Font.Size = 20
        .Font.Bold = True
    End With
    mRow = mRow + 1
    Body ws, "Where every result comes from. SOURCES (top) lists each state's official public map and the data " & _
        "service this tool reads; QUIRKS & CAVEATS (bottom) explains, in plain language, the things about that " & _
        "data that shape the answer. This tool is a screening aid, not an authoritative source - read the " & _
        "disclaimer on Start Here and verify every point on the official map."
    mRow = mRow + 1

    ' ===================== SECTION 1: SOURCES =====================
    Section ws, "1.  SOURCES"
    Body ws, "For each state: the OFFICIAL PUBLIC MAP (open it to confirm a point yourself) and the DATA SERVICE " & _
        "this tool reads. Michigan, Indiana and Wisconsin are wired for road classification. Minnesota, Illinois " & _
        "and Ohio are reference-only for now - the tool runs the urban-boundary check on them but does not yet " & _
        "read their class layer; their sources are listed so the data can be wired in later."
    mRow = mRow + 1

    StateHeader ws, "Michigan (MDOT) - WIRED"
    Cite ws, "Michigan Department of Transportation, National Functional Classification, NHS & ACUB (interactive map),", APP_MI
    Cite ws, "Michigan Department of Transportation, Functional System, ArcGIS REST feature service layer 353 " & _
        "(NextGen PR Finder; road names from companion layer 543) - queried by this tool,", REST_MDOT_NFC

    StateHeader ws, "Indiana (INDOT) - WIRED"
    Cite ws, "Indiana Department of Transportation, INDOT Functional Class Map (Functional Classification & Urban " & _
        "Area Boundary),", APP_IN
    Cite ws, "Indiana Department of Transportation, LRSE_Functional_Class, ArcGIS REST feature service layer 22 " & _
        "(road names from the 2021 centerlines layer) - queried by this tool,", REST_IN_NFC

    StateHeader ws, "Wisconsin (WisDOT) - WIRED"
    Cite ws, "Wisconsin Department of Transportation, Functional Classification (official page; static county and " & _
        "urban-area PDF maps - WisDOT publishes no statewide interactive map),", APP_WI
    Cite ws, "Wisconsin Department of Transportation, Functional Class - State Trunk Network, ArcGIS REST feature " & _
        "service layer 3 (local roads from the Flood Damage Assessment snapshot, layer 1) - queried by this tool,", REST_WI_STATE_TRUNK

    StateHeader ws, "Minnesota (MnDOT) - reference only, not yet wired"
    Cite ws, "Minnesota Department of Transportation, Enterprise MnDOT Mapping Application (EMMA) - Functional " & _
        "Class layer,", APP_MN
    Cite ws, "Minnesota Department of Transportation, Functional Class, ArcGIS REST map service layer 11 " & _
        "(mndot_commonlayers2),", REST_MN_NFC

    StateHeader ws, "Illinois (IDOT) - reference only, not yet wired"
    Cite ws, "Illinois Department of Transportation, Roadway Functional Class (Getting Around Illinois map viewer),", APP_IL
    Cite ws, "Illinois Department of Transportation, Functional Class, ArcGIS REST map service layer 0,", REST_IL_NFC

    StateHeader ws, "Ohio (ODOT) - reference only, not yet wired"
    Cite ws, "Ohio Department of Transportation, Transportation Information Mapping System (TIMS),", APP_OH
    Cite ws, "Ohio Department of Transportation, Functional Class, ArcGIS REST map service layer 0,", REST_OH_NFC

    StateHeader ws, "All states - urban boundary, street names, geocoding, flood maps"
    Cite ws, "U.S. Department of Transportation, Bureau of Transportation Statistics, 2020 Adjusted Urban Area " & _
        "Boundaries (National Transportation Atlas Database) - the urban/rural source for every state,", REST_ACUB
    Cite ws, "U.S. Census Bureau, TIGERweb Transportation - Local Roads, ArcGIS REST map service layer 8 " & _
        "(street names, all states),", REST_TIGER_ROADS
    Cite ws, "U.S. Census Bureau, Census Geocoder - one-line address service (fills coordinates from an address),", REST_CENSUS_GEOCODE
    Cite ws, "Federal Emergency Management Agency, Map Service Center - Print FIRMette service (flood-map extracts),", REST_FIRMETTE
    mRow = mRow + 1

    ' ===================== SECTION 2: QUIRKS & CAVEATS =====================
    Section ws, "2.  QUIRKS & CAVEATS (plain language)"

    Sub2 ws, "This is a screening aid, not a determination"
    Body ws, "This tool does NOT authoritatively identify FHWA federal-aid roads. It flags high-probability " & _
        "candidates for a person to review, and can miss or mis-tag a road. Every coordinate must be verified by " & _
        "a human on the official map above before it is relied on. Results are informational only and do NOT " & _
        "constitute a federal-aid, funding, or eligibility determination."

    Sub2 ws, "How the answer is built"
    Body ws, "For each point the tool reads the road's functional class from the state's own data, and separately " & _
        "checks whether the point falls inside a Census-adjusted urban area. A road that is Urban Minor Collector " & _
        "or higher is tagged 'federal aid'; Rural Local, Urban Local and Rural Minor Collector are 'non-federal " & _
        "aid'. It tags the road, not the project - what that means for a work order is the reviewer's call."

    Sub2 ws, "BOUNDARY ROADS (urban vs rural on the edge)"
    Body ws, "Urban vs rural comes only from the urban-boundary layer, never from the state's own data, so every " & _
        "state is judged the same way. When a GPS point sits ON or just OUTSIDE an urban boundary - e.g. on a " & _
        "road that forms the boundary, or a few feet onto the rural side - the point is deliberately treated as " & _
        "URBAN (the boundary check always searches at least 200 ft). This leans toward flagging a boundary road " & _
        "for review rather than silently dropping it. Always confirm boundary cases manually on the source map."

    Sub2 ws, "Michigan"
    Body ws, "Michigan's road data includes retired (historical) segments; the tool skips those so it never reads " & _
        "an out-of-date classification. Road names exist only for numbered state routes (like 'I-94'), so local " & _
        "street names are filled in from the Census street layer instead."

    Sub2 ws, "Indiana"
    Body ws, "Indiana marks each record's status; the tool uses only Active records. Indiana's class data carries " & _
        "no road name at all, so names come from a separate centerline layer and the Census - expect the Road " & _
        "Name column to be blank more often for Indiana."

    Sub2 ws, "Wisconsin"
    Body ws, "Wisconsin data comes in two layers: state highways are checked first, then local (county/city/town) " & _
        "roads. The local-roads layer bakes urban/rural into its own class number, which the tool translates back " & _
        "to a plain federal class before deciding. One local code ('Urban Collector') doesn't separate major from " & _
        "minor collector - it's treated as major, which never changes the answer because every urban collector is " & _
        "federal aid regardless."

    Sub2 ws, "Minnesota, Illinois, Ohio (not yet wired)"
    Body ws, "Road-class lookup is not wired for these states yet, so the tool runs only the urban-boundary check " & _
        "and the Federal Aid Status column says 'ACUB only - class lookup not wired for this state' rather than " & _
        "guessing. The sources listed above are where each state's class data lives when it is added."

    Sub2 ws, "Street names and addresses (all states)"
    Body ws, "Street names are backfilled from the U.S. Census street layer wherever the state data only knows " & _
        "numbered routes. If a row has an address but no coordinates, Check Roads geocodes it with the free Census " & _
        "address service first (it never overwrites coordinates you typed in)."
    mRow = mRow + 1

    Body ws, "Sources verified 2026-07-05 against each service's live metadata. This tool classifies the road, " & _
        "never the project.  " & ProductTitle() & " - " & BUILD_REFERENCE & "."

    HideGridlines ws
    ws.Tab.Color = RGB(120, 120, 120)
End Sub

Private Sub Section(ByVal ws As Worksheet, ByVal txt As String)
    With ws.Cells(mRow, 2)
        .Value = txt
        .Font.Bold = True
        .Font.Size = 15
        .Font.Color = RGB(47, 79, 79)
    End With
    mRow = mRow + 1
End Sub

Private Sub StateHeader(ByVal ws As Worksheet, ByVal txt As String)
    With ws.Cells(mRow, 2)
        .Value = txt
        .Font.Bold = True
        .Font.Size = 12
    End With
    mRow = mRow + 1
End Sub

Private Sub Sub2(ByVal ws As Worksheet, ByVal txt As String)
    With ws.Cells(mRow, 2)
        .Value = txt
        .Font.Bold = True
        .Font.Size = 12
    End With
    mRow = mRow + 1
End Sub

Private Sub Body(ByVal ws As Worksheet, ByVal txt As String)
    With ws.Cells(mRow, 2)
        .Value = txt
        .WrapText = True
        .Font.Color = RGB(60, 60, 60)
        .VerticalAlignment = xlTop
    End With
    ws.Rows(mRow).AutoFit
    mRow = mRow + 1
End Sub

' One bluebook-style citation: the source sentence, then "available at: <url>",
' with the whole line hyperlinked to the URL so it's one click to the source.
Private Sub Cite(ByVal ws As Worksheet, ByVal citation As String, ByVal url As String)
    Dim cell As Range
    Set cell = ws.Cells(mRow, 2)
    ws.Hyperlinks.Add Anchor:=cell, Address:=url, TextToDisplay:="    " & citation & " available at: " & url
    cell.WrapText = True
    cell.VerticalAlignment = xlTop
    cell.Font.Size = 10
    ws.Rows(mRow).AutoFit
    mRow = mRow + 1
End Sub
