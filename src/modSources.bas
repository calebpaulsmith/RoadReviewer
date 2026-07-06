Attribute VB_Name = "modSources"
Option Explicit

' Sources sheet - per-state data-source citations and every schema quirk a
' reviewer might need to defend a verdict. Static text written at build
' time (CLAUDE.md §4.2/§4.2a/§4.2b are the authoritative research notes;
' keep this sheet in sync with them and with web/sources.html).

Private mRow As Long   ' running row cursor while writing the sheet

Public Sub BuildSourcesSheet()
    Dim ws As Worksheet
    Set ws = FreshSheet(SH_SOURCES)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 120
    mRow = 2

    With ws.Cells(mRow, 2)
        .Value = "Data Sources & Quirks"
        .Font.Size = 20
        .Font.Bold = True
    End With
    mRow = mRow + 1
    Body ws, "Every service this tool queries, with the exact layer, the fields read, and the quirks that shape the results."
    mRow = mRow + 1

    Section ws, "Read this first - what this tool is (and isn't)"
    Body ws, "This tool does NOT authoritatively identify FHWA federal-aid roads. It flags high-probability " & _
        "candidates for a person to review, and can miss or mis-tag a road. It is not an authoritative source for " & _
        "FHWA functional classification."
    Body ws, "Every coordinate must be verified by a human against the official source map before it is relied on - " & _
        "use each Sites row's NFC Map link and the citations below. Results are informational only and do NOT " & _
        "constitute a federal-aid, funding, or eligibility determination."
    mRow = mRow + 1

    Section ws, "How the federal-aid check works"
    Body ws, "The FHWA functional class comes from each state's own roads layer. Urban vs Rural comes from the " & _
        "nationwide 2020 Adjusted Urban Boundary (ACUB) polygons - never from the state layer, even when it " & _
        "carries its own urban flag, so every state is judged the same way."
    Body ws, "A road is tagged FEDERAL AID when any segment at the point (or within the search buffer) is Urban " & _
        "Minor Collector or greater. Non-federal-aid classes: Rural Local, Urban Local, Rural Minor Collector. " & _
        "The tool classifies the road, never the project - the reviewer decides what federal-aid status means " & _
        "for a given work order."
    Body ws, "Every lookup tries the exact point first, then retries with a fallback radius (the Search buffer, " & _
        "default 200 ft)."
    Body ws, "BOUNDARY ROADS: the urban-boundary check always uses a radius of at least 200 ft. So when a GPS point " & _
        "sits ON or just OUTSIDE an urban boundary - e.g. on a road that itself forms the boundary, or a few feet " & _
        "onto the rural side - the point is deliberately treated as URBAN (which leans toward a federal-aid flag " & _
        "for review) rather than being missed. This is intentional: it is safer to over-flag a boundary road for " & _
        "human review than to silently drop it. Always confirm boundary cases manually on the source map."
    mRow = mRow + 1

    Section ws, "Michigan - MDOT"
    Body ws, "Functional class: NextGen PR Finder, layer 353 ""Functional System"" - field FunctionalSystem " & _
        "(FHWA codes 1-7), filtered to RHRetireDate IS NULL so retired historical segments are excluded."
    LinkLine ws, "MDOT layer 353 (Functional System)", REST_MDOT_NFC
    Body ws, "Route names: companion layer 543 ""Route System"" (RouteDesignation + RouteNumber). Only designated " & _
        "trunkline routes (I-, US-, M-) are populated - local street names come from the Census TIGER layer below."
    LinkLine ws, "MDOT layer 543 (Route System)", REST_MDOT_ROUTE
    Body ws, "Quirk: MDOT's server returns HTTP 403 to requests without a browser User-Agent; this tool sends a " & _
        "browser-style UA on every request."
    mRow = mRow + 1

    Section ws, "Indiana - INDOT / IndianaMap"
    Body ws, "Functional class: LRSE_Functional_Class, layer 22 - field functional_class (FHWA codes 1-7), " & _
        "filtered to record_status = 5 (Active), Indiana's analog to Michigan's retire-date filter."
    LinkLine ws, "INDOT LRSE_Functional_Class layer 22", REST_IN_NFC
    Body ws, "Road name: Road Centerlines of Indiana 2021, layer 15 - field st_full. This layer shares no join key " & _
        "with the class layer, so it is a separate point query; the name is blank more often than Michigan's, " & _
        "and the Census TIGER street-name column backs it up."
    LinkLine ws, "Indiana road centerlines layer 15", REST_IN_ROADNAME
    mRow = mRow + 1

    Section ws, "Wisconsin - WisDOT"
    Body ws, "Two layers, queried in sequence. First the State Trunk Network extract (interstates / US and state " & _
        "highways) - field FED_FC_CD is already a bare FHWA 1-7 code; road name is built from HWYTYPE + HWYNUM + " & _
        "HWYDIR (e.g. ""STH 32 S"")."
    LinkLine ws, "WisDOT State Trunk Network (FFCL_gdb layer 3)", REST_WI_STATE_TRUNK
    Body ws, "Only when no state-trunk segment intersects: the Local Road Network flood-damage snapshot - field " & _
        "FNCT_CLS_CTGY_TYCD, WisDOT's own code that embeds Urban/Rural into the number (45 = Rural Local, 97 = " & _
        "Urban Local, ...). The tool strips that back to a bare FHWA code; ACUB stays the urban/rural source of truth."
    LinkLine ws, "WisDOT Local Road Network snapshot (layer 1)", REST_WI_LOCAL_ROADS
    Body ws, "Quirk: local-roads code 96 ""Urban Collector Other"" doesn't distinguish Major from Minor Collector. " & _
        "It is mapped to Major Collector, which can never change the verdict - every URBAN collector is federal " & _
        "aid regardless of the major/minor split."
    mRow = mRow + 1

    Section ws, "Urban / Rural - ACUB (all states)"
    Body ws, "USDOT BTS National Transportation Atlas Database, ""2020 Adjusted Urban Area Boundaries"" (nationwide " & _
        "polygon layer; fields NAME and UACE). A point inside any polygon is Urban; otherwise Rural. This is the " & _
        "single urban/rural source of truth for every state."
    LinkLine ws, "NTAD 2020 Adjusted Urban Areas", REST_ACUB
    mRow = mRow + 1

    Section ws, "Street names - U.S. Census Bureau TIGERweb (all states)"
    Body ws, "Local Roads layer (Transportation MapServer, layer 8), field NAME. Fills the Street Name column for " & _
        "every state, covering the local streets the state layers name poorly or not at all. Multiple streets " & _
        "inside the search buffer (intersections) are pipe-joined."
    LinkLine ws, "TIGERweb Local Roads layer 8", REST_TIGER_ROADS
    mRow = mRow + 1

    Section ws, "Address geocoding - U.S. Census Bureau (all states)"
    Body ws, "The Census one-line address geocoder runs automatically during Check Roads for any row that has an " & _
        "Address but no coordinates. Coordinates already typed into a row are never overwritten. Free, no " & _
        "account, US addresses only."
    LinkLine ws, "Census one-line address geocoder", REST_CENSUS_GEOCODE
    mRow = mRow + 1

    Section ws, "Flood maps - FEMA Map Service Center"
    Body ws, "Each row's FIRMette Portal link opens FEMA's Map Service Center FIRMette page at the point."
    If ProductIsInspector() Then
        Body ws, "The Download FIRMettes button drives FEMA's Print FIRMette geoprocessing service (submit job, " & _
            "poll status, download the rendered PDF) for every row."
        LinkLine ws, "FEMA Print FIRMette GP service", REST_FIRMETTE
    End If
    mRow = mRow + 1

    Section ws, "States not yet wired (MN / IL / OH)"
    Body ws, "No functional-class layer is wired for these states yet. The ACUB urban/rural check still runs on " & _
        "every row; the Federal Aid Status column reads ""ACUB only - class lookup not wired for this state"" " & _
        "so nothing silently pretends to be classified."
    mRow = mRow + 1

    Body ws, "All layer names, fields and quirks were confirmed against each service's live REST metadata " & _
        "(?f=json), 2026-07-01 through 2026-07-05. This tool classifies the road, never the project. " & _
        ProductTitle() & " - " & BUILD_REFERENCE & "."

    HideGridlines ws
    ws.Tab.Color = RGB(120, 120, 120)
End Sub

Private Sub Section(ByVal ws As Worksheet, ByVal txt As String)
    With ws.Cells(mRow, 2)
        .Value = txt
        .Font.Bold = True
        .Font.Size = 13
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

Private Sub LinkLine(ByVal ws As Worksheet, ByVal label As String, ByVal url As String)
    Dim cell As Range
    Set cell = ws.Cells(mRow, 2)
    ws.Hyperlinks.Add Anchor:=cell, Address:=url, TextToDisplay:="    " & label & "  ->  " & url
    cell.Font.Size = 10
    mRow = mRow + 1
End Sub
