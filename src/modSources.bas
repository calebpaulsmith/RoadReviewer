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

' Official public functional-class app URLs (APP_MI..APP_OH) now live in
' modConstants (they also drive the Open column); used below in the citations.

' Reference-only NFC class services for the not-yet-wired states now live in
' modConstants (public REST_MN_NFC/REST_IL_NFC/REST_OH_NFC) so ServiceDefault
' and the Svc_ override table can resolve them too.

Public Sub BuildSourcesSheet()
    Dim ws As Worksheet
    Set ws = FreshSheet(SH_SOURCES)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 120
    mRow = 2

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
    Cite ws, "Wisconsin Department of Transportation, Functional Class - Wisconsin Local Road Network, ArcGIS REST " & _
        "feature service layer 1 (queried FIRST - local roads and most collectors),", REST_WI_LOCAL_ROADS
    Cite ws, "Wisconsin Department of Transportation, Functional Class - State Trunk Network, ArcGIS REST feature " & _
        "service layer 3 (state highways and interstates; queried when the local layer has no class),", REST_WI_STATE_TRUNK

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
    Body ws, "Wisconsin data comes in two layers. The LOCAL road layer (county/city/town roads and most collectors) " & _
        "is checked FIRST because it covers most points; the STATE TRUNK layer (interstates and state highways) is " & _
        "checked only when the local layer has no class for the point. State highways also appear in the local " & _
        "layer but with no class - the tool treats those as 'not found here' and falls through to the trunk layer. " & _
        "The local layer bakes urban/rural into its own class number, which the tool translates back to a plain " & _
        "federal class. One local code ('Urban Collector') doesn't separate major from minor collector - it's " & _
        "treated as major, which never changes the answer because every urban collector is federal aid regardless. " & _
        "Note: WisDOT moved its local-roads service in 2026 (the old one now needs a login); if it moves again, " & _
        "paste the new layer URL into the Service URLs table below - no rebuild needed."

    Sub2 ws, "Minnesota, Illinois, Ohio (not yet wired)"
    Body ws, "Road-class lookup is not wired for these states yet, so the tool runs only the urban-boundary check " & _
        "and the Federal Aid Status column says 'ACUB only - class lookup not wired for this state' rather than " & _
        "guessing. The sources listed above are where each state's class data lives when it is added."

    Sub2 ws, "Street names and addresses (all states)"
    Body ws, "Street names are backfilled from the U.S. Census street layer wherever the state data only knows " & _
        "numbered routes. If a row has an address but no coordinates, Check Roads geocodes it with the free Census " & _
        "address service first (it never overwrites coordinates you typed in)."
    mRow = mRow + 1

    ' ===================== SECTION 3: SERVICE URLs (overrides) =====================
    Section ws, "3.  SERVICE URLs  (paste a replacement to swap a layer)"
    Body ws, "Each row below is a live data service this tool reads. Leave a cell BLANK to use the built-in default " & _
        "shown next to it. To re-point a layer a state has moved or locked (the way WisDOT just moved its " & _
        "local-roads layer), paste the full replacement layer URL - e.g. https://services5.arcgis.com/.../" & _
        "FeatureServer/1 - into the yellow cell in column C. It takes effect on the next Check Roads, no rebuild " & _
        "needed. The same short KEY names are used by the RoadReviewer website and the AGOL notebook, so one URL " & _
        "swap is easy to mirror across every product. Tip: open the default URL in a browser and add '?f=pjson' " & _
        "to confirm a replacement responds before you trust it."
    mRow = mRow + 1

    ws.Columns("C").ColumnWidth = 90
    SvcHeaderRow ws
    SvcOverrideRow ws, "WI_LOCAL_ROADS", "Wisconsin - local roads + most collectors (queried first)"
    SvcOverrideRow ws, "WI_STATE_TRUNK", "Wisconsin - state highways & interstates (fallback)"
    SvcOverrideRow ws, "MI_NFC", "Michigan - functional class (layer 353)"
    SvcOverrideRow ws, "MI_ROUTE", "Michigan - trunkline route names (layer 543)"
    SvcOverrideRow ws, "IN_NFC", "Indiana - functional class (LRSE, layer 22)"
    SvcOverrideRow ws, "IN_ROADNAME", "Indiana - road names (2021 centerlines)"
    SvcOverrideRow ws, "ACUB", "All states - adjusted urban-area boundary (urban/rural)"
    SvcOverrideRow ws, "TIGER_ROADS", "All states - Census street names"
    SvcOverrideRow ws, "MN_NFC", "Minnesota - functional class (reference; not yet wired)"
    SvcOverrideRow ws, "IL_NFC", "Illinois - functional class (reference; not yet wired)"
    SvcOverrideRow ws, "OH_NFC", "Ohio - functional class (reference; not yet wired)"
    mRow = mRow + 1

    Body ws, "Sources verified 2026-07-05 against each service's live metadata. This tool classifies the road, " & _
        "never the project.  " & ProductTitle() & " - " & BUILD_REFERENCE & "."

    mRow = mRow + 1
    Dim credit As Range
    Set credit = ws.Cells(mRow, 2)
    ws.Hyperlinks.Add Anchor:=credit, Address:="mailto:caleb.smith@fema.dhs.gov", _
        TextToDisplay:="Created by Caleb Smith. Reach out to caleb.smith@fema.dhs.gov for any questions."
    credit.Font.Size = 11
    credit.Font.Bold = True
    mRow = mRow + 1

    HideGridlines ws
    ws.Tab.Color = RGB(120, 120, 120)
End Sub

' Header row for the Service URLs override table.
Private Sub SvcHeaderRow(ByVal ws As Worksheet)
    ws.Cells(mRow, 2).Value = "KEY / layer (built-in default)"
    ws.Cells(mRow, 3).Value = "Paste replacement URL here (blank = use default)"
    ws.Range(ws.Cells(mRow, 2), ws.Cells(mRow, 3)).Font.Bold = True
    mRow = mRow + 1
End Sub

' One override row: label + default URL in col B, an editable (named-range)
' paste cell in col C tinted like an input field. The named range Svc_<key> is
' what ServiceUrl(key) reads at query time.
Private Sub SvcOverrideRow(ByVal ws As Worksheet, ByVal key As String, ByVal label As String)
    With ws.Cells(mRow, 2)
        .Value = key & "  -  " & label & "   (default: " & ServiceDefault(key) & ")"
        .WrapText = True
        .Font.Size = 9
        .Font.Color = RGB(60, 60, 60)
        .VerticalAlignment = xlTop
    End With
    Dim cell As Range
    Set cell = ws.Cells(mRow, 3)
    cell.Interior.Color = RGB(255, 255, 204)      ' input yellow
    cell.Borders.Color = RGB(200, 200, 160)
    cell.WrapText = True
    cell.VerticalAlignment = xlTop
    AddNameForCell cell, "Svc_" & key
    ws.Rows(mRow).AutoFit
    mRow = mRow + 1
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
