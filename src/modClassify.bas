Attribute VB_Name = "modClassify"
Option Explicit

' RoadReviewer V1 - Workflow 1: Classify Roads (§4.2 query strategy, F7 rule).
' For each Sites row: query the state's NFC layer (class + road name) and the
' nationwide NTAD ACUB layer (urban/rural), then write the eligibility verdict
' back to the row. Michigan, Indiana and Wisconsin are wired in V1; other
' states still get the ACUB check (F8).

Private Const MI_NFC_OUTFIELDS As String = "FunctionalSystem,PR"
Private Const MI_ROUTE_OUTFIELDS As String = "RouteDesignation,RouteNumber"
Private Const IN_NFC_OUTFIELDS As String = "functional_class"
Private Const IN_ROADNAME_OUTFIELDS As String = "st_full"
Private Const WI_TRUNK_OUTFIELDS As String = "FED_FC_CD,HWYTYPE,HWYNUM,HWYDIR"
Private Const WI_LOCAL_OUTFIELDS As String = "FNCT_CLS_CTGY_TYCD,ST_LABL_NM"
Private Const ACUB_OUTFIELDS As String = "NAME,UACE,state_1"

' The ACUB (urban-boundary) fallback search always uses at least this many
' feet, even if JobBufferFeet has been narrowed below it for precise road
' matching in a dense urban grid. Without this floor, narrowing the shared
' buffer for road-matching purposes would silently widen the "point right on
' the edge of the urban boundary gets flagged Rural" bug this exists to
' prevent - a site sitting a few feet on the wrong side of a road that
' itself touches the boundary should still resolve Urban.
Private Const ACUB_MIN_BUFFER_FEET As Long = 200

Public Sub ClassifyAllRows()
    ClassifyRows False
End Sub

Public Sub ReRunFailedClassifications()
    ClassifyRows True
End Sub

Private Sub ClassifyRows(ByVal onlyFailed As Boolean)
    Dim ws As Worksheet, last As Long, r As Long
    Dim stateCode As String, stateHasNfc As Boolean
    Dim processed As Long, total As Long

    Set ws = SitesSheet()
    last = SitesLastRow()
    If last < SITES_FIRST_DATA_ROW Then
        If Not gHeadless Then MsgBox "No site rows found. Add points on the Sites sheet first.", vbInformation, "Classify Roads"
        Exit Sub
    End If

    stateCode = UCase$(SetupValue(NR_STATE))
    If Len(stateCode) = 0 Then stateCode = "MI"
    ' Local var deliberately NOT named NfcWired - VBA identifiers are
    ' case-insensitive, so a local variable with the same name as the
    ' module-level NfcWired() function shadows it, and `NfcWired(stateCode)`
    ' on the right-hand side gets parsed as indexing a Boolean scalar like an
    ' array ("Expected array" compile error) instead of calling the function.
    stateHasNfc = NfcWired(stateCode)
    If Not stateHasNfc Then
        If Not gHeadless Then MsgBox "Road-class (NFC) lookup is not yet wired for " & stateCode & "." & vbCrLf & _
            "The ACUB urban-boundary check will still run on every row.", vbInformation, "Classify Roads"
    End If

    For r = SITES_FIRST_DATA_ROW To last
        If Not RowIsEmpty(ws, r) Then
            If (Not onlyFailed) Or RowIsFailed(ws, r) Then total = total + 1
        End If
    Next r
    If total = 0 Then
        If Not gHeadless Then MsgBox IIf(onlyFailed, "No failed rows to re-run.", "No site rows to classify."), vbInformation, "Classify Roads"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo Done
    For r = SITES_FIRST_DATA_ROW To last
        If RowIsEmpty(ws, r) Then GoTo NextRow
        If onlyFailed And Not RowIsFailed(ws, r) Then GoTo NextRow
        processed = processed + 1
        SetStatus "Classifying " & processed & " of " & total & " - " & _
            CStr(ws.Cells(r, COL_SITENAME).Value)
        ClassifyOneRow ws, r, stateCode
        DoEvents
NextRow:
    Next r

Done:
    Application.ScreenUpdating = True
    ClearStatus
    If Err.Number <> 0 Then
        If Not gHeadless Then MsgBox "Classification stopped: " & Err.Description, vbExclamation, "Classify Roads"
    Else
        If Not gHeadless Then MsgBox "Classified " & processed & " row(s).", vbInformation, "Classify Roads"
    End If
End Sub

Private Function RowIsFailed(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    RowIsFailed = (InStr(1, CStr(ws.Cells(r, COL_ELIGIBILITY).Value), STATUS_FAILED_PREFIX, vbTextCompare) = 1)
End Function

' States whose NFC (road class) layer is wired in V1. Every state still gets
' the ACUB urban-boundary check regardless (F8).
Private Function NfcWired(ByVal stateCode As String) As Boolean
    Select Case stateCode
        Case "MI", "IN", "WI": NfcWired = True
        Case Else: NfcWired = False
    End Select
End Function

Private Sub ClassifyOneRow(ByVal ws As Worksheet, ByVal r As Long, ByVal stateCode As String)
    TraceLine "ClassifyOneRow row=" & r & " name=" & CStr(ws.Cells(r, COL_SITENAME).Value)
    ' Idempotent: clear prior lookup output before writing (N5).
    ClearLookupCells ws, r

    If Not HasValidCoords(ws, r) Then
        ws.Cells(r, COL_ELIGIBILITY).Value = STATUS_FAILED_PREFIX & "no/invalid coordinates"
        TraceLine "  no/invalid coords"
        Exit Sub
    End If

    Dim lat As String, lon As String, errMsg As String
    lat = InvariantNum(ws.Cells(r, COL_LAT).Value)
    lon = InvariantNum(ws.Cells(r, COL_LON).Value)

    ' --- ACUB (urban/rural) - runs for every state, and is the single
    ' source of truth for the Urban/Rural column even on states whose own
    ' NFC layer happens to carry its own urban/rural flag (§4.2). Uses its
    ' own floor-guaranteed buffer (AcubBufferFeet), not the raw road-search
    ' BufferFeet, so a point sitting just outside an ACUB polygon - e.g. on
    ' the wrong side of a road that itself touches the boundary - still
    ' resolves Urban even if JobBufferFeet has been narrowed for precise
    ' road matching in a dense grid. ---
    Dim acubJson As String, isUrban As Boolean, acubName As String
    acubJson = QueryWithFallback(REST_ACUB, lat, lon, ACUB_OUTFIELDS, "1=1", AcubBufferFeet(), errMsg)
    If Len(errMsg) > 0 Then
        ws.Cells(r, COL_ELIGIBILITY).Value = STATUS_FAILED_PREFIX & "ACUB query (" & errMsg & ")"
        Exit Sub
    End If
    isUrban = (FeatureCount(acubJson) > 0)
    If isUrban Then acubName = FirstString(acubJson, "NAME")
    ws.Cells(r, COL_URBANRURAL).Value = IIf(isUrban, "Urban", "Rural")
    ws.Cells(r, COL_ACUBNAME).Value = acubName

    If Not NfcWired(stateCode) Then
        ws.Cells(r, COL_ELIGIBILITY).Value = "ACUB only - class lookup not wired for this state"
        Exit Sub
    End If

    ' --- NFC class + road name (state-specific layer) ---
    Dim codes As Collection, roadName As String
    Select Case stateCode
        Case "MI": QueryMichiganNfc lat, lon, codes, roadName, errMsg
        Case "IN": QueryIndianaNfc lat, lon, codes, roadName, errMsg
        Case "WI": QueryWisconsinNfc lat, lon, codes, roadName, errMsg
    End Select
    If Len(errMsg) > 0 Then
        ws.Cells(r, COL_ELIGIBILITY).Value = STATUS_FAILED_PREFIX & "NFC query (" & errMsg & ")"
        Exit Sub
    End If
    ws.Cells(r, COL_ROADNAME).Value = roadName

    ' --- Street name(s) via Census TIGER (covers everything the state's own
    ' road-name field misses — e.g. MDOT 543 and Indiana's centerline layer
    ' both only reliably carry designated/trunkline names). Failures here
    ' are non-fatal — keep going so the classification still lands even if
    ' TIGER is briefly unavailable. ---
    Dim tigerErr As String, streets As String
    streets = LookupStreetNames(lat, lon, tigerErr)
    ws.Cells(r, COL_STREET).Value = streets

    ' --- Class label + Federal Aid Status verdict ---
    ws.Cells(r, COL_CLASS).Value = ClassLabels(codes)
    ws.Cells(r, COL_ELIGIBILITY).Value = FederalAidVerdict(codes, isUrban)
End Sub

' Michigan: NFC class from MDOT layer 353 (FunctionalSystem, bare FHWA 1-7,
' §4.2), road name from the companion route-designation layer 543 (blank for
' local streets - that's expected, TIGER fills the gap). Both filtered to
' exclude retired LRS segments.
Private Sub QueryMichiganNfc(ByVal lat As String, ByVal lon As String, _
        ByRef codes As Collection, ByRef roadName As String, ByRef errMsg As String)
    Dim nfcJson As String, routeJson As String, routeErr As String
    nfcJson = QueryWithFallback(REST_MDOT_NFC, lat, lon, MI_NFC_OUTFIELDS, "RHRetireDate IS NULL", BufferFeet(), errMsg)
    If Len(errMsg) > 0 Then Exit Sub
    Set codes = ExtractIntegers(nfcJson, "FunctionalSystem")

    routeJson = QueryWithFallback(REST_MDOT_ROUTE, lat, lon, MI_ROUTE_OUTFIELDS, "RHRetireDate IS NULL", BufferFeet(), routeErr)
    If Len(routeErr) = 0 Then roadName = BuildRouteName(routeJson)
End Sub

' Indiana: NFC class from LRSE_Functional_Class layer 22 (functional_class,
' bare FHWA 1-7, no urban/rural embedded - §4.2a). record_status=5 (Active)
' is Indiana's analog to Michigan's RHRetireDate filter. Road name is a
' separate, un-keyed point query against the statewide centerline layer -
' Indiana's functional-class layer carries no name field at all (not even a
' trunkline-only one like MDOT 543), so this is blank more often; TIGER
' backs it up.
Private Sub QueryIndianaNfc(ByVal lat As String, ByVal lon As String, _
        ByRef codes As Collection, ByRef roadName As String, ByRef errMsg As String)
    Dim nfcJson As String, nameJson As String, nameErr As String
    nfcJson = QueryWithFallback(REST_IN_NFC, lat, lon, IN_NFC_OUTFIELDS, "record_status=5", BufferFeet(), errMsg)
    If Len(errMsg) > 0 Then Exit Sub
    Set codes = ExtractIntegers(nfcJson, "functional_class")

    nameJson = QueryWithFallback(REST_IN_ROADNAME, lat, lon, IN_ROADNAME_OUTFIELDS, "1=1", BufferFeet(), nameErr)
    If Len(nameErr) = 0 Then roadName = FirstString(nameJson, "st_full")
End Sub

' Wisconsin: two layers queried in sequence (§4.2b). Try the State Trunk
' Network first (FED_FC_CD is already a bare FHWA 1-7 string) since it's the
' more authoritative/current source for state highways. Only when nothing
' intersects there does it fall back to the Local Road Network snapshot,
' whose FNCT_CLS_CTGY_TYCD needs WisconsinLocalCategoryToFhwa() to strip the
' embedded urban/rural digit back out to a bare FHWA code. Neither layer
' needs a retired-segment filter (both are point-in-time snapshot extracts,
' confirmed via live schema probe) or a browser User-Agent (AGOL-hosted,
' same pattern as the nationwide ACUB layer).
Private Sub QueryWisconsinNfc(ByVal lat As String, ByVal lon As String, _
        ByRef codes As Collection, ByRef roadName As String, ByRef errMsg As String)
    Dim trunkJson As String, localJson As String, fc As Variant, cat As Variant
    Set codes = New Collection

    trunkJson = QueryWithFallback(REST_WI_STATE_TRUNK, lat, lon, WI_TRUNK_OUTFIELDS, "1=1", BufferFeet(), errMsg)
    If Len(errMsg) > 0 Then Exit Sub

    If FeatureCount(trunkJson) > 0 Then
        For Each fc In ExtractStrings(trunkJson, "FED_FC_CD")
            If IsNumeric(fc) Then codes.Add CLng(fc)
        Next fc
        roadName = BuildWisconsinTrunkRoadName(trunkJson)
        Exit Sub
    End If

    localJson = QueryWithFallback(REST_WI_LOCAL_ROADS, lat, lon, WI_LOCAL_OUTFIELDS, "1=1", BufferFeet(), errMsg)
    If Len(errMsg) > 0 Then Exit Sub
    For Each cat In ExtractIntegers(localJson, "FNCT_CLS_CTGY_TYCD")
        codes.Add WisconsinLocalCategoryToFhwa(CLng(cat))
    Next cat
    roadName = FirstString(localJson, "ST_LABL_NM")
End Sub

Private Function BuildWisconsinTrunkRoadName(ByVal trunkJson As String) As String
    Dim hwytype As String, hwynum As String, hwydir As String
    hwytype = FirstString(trunkJson, "HWYTYPE")
    hwynum = FirstString(trunkJson, "HWYNUM")
    hwydir = FirstString(trunkJson, "HWYDIR")
    BuildWisconsinTrunkRoadName = Trim$(hwytype & " " & hwynum & " " & hwydir)
End Function

' Query Census TIGER Local Roads (layer 8) within the search buffer and
' return a pipe-joined list of unique street names. Returns "" with errMsg
' set if the call failed; "" with errMsg empty if no streets matched.
Private Function LookupStreetNames(ByVal lat As String, ByVal lon As String, _
        ByRef errMsg As String) As String
    Dim json As String, names As Collection, seen As String, out As String, nm As Variant
    json = RunQuery(REST_TIGER_ROADS, lat, lon, "NAME", "1=1", BufferFeet(), errMsg)
    If Len(errMsg) > 0 Then Exit Function
    Set names = ExtractStrings(json, "NAME")
    For Each nm In names
        If Len(Trim$(CStr(nm))) > 0 And InStr(seen, "|" & CStr(nm) & "|") = 0 Then
            seen = seen & "|" & CStr(nm) & "|"
            out = out & IIf(Len(out) > 0, " | ", "") & CStr(nm)
        End If
    Next nm
    LookupStreetNames = out
End Function

Private Sub ClearLookupCells(ByVal ws As Worksheet, ByVal r As Long)
    ws.Cells(r, COL_CLASS).ClearContents
    ws.Cells(r, COL_STREET).ClearContents
    ws.Cells(r, COL_URBANRURAL).ClearContents
    ws.Cells(r, COL_ACUBNAME).ClearContents
    ws.Cells(r, COL_ROADNAME).ClearContents
    ws.Cells(r, COL_ELIGIBILITY).ClearContents
End Sub

' Exact point intersect first; if no hit, retry with the given buffer (§4.2).
Private Function QueryWithFallback(ByVal baseUrl As String, ByVal lat As String, ByVal lon As String, _
        ByVal outFields As String, ByVal whereClause As String, ByVal fallbackDistanceFt As Long, _
        ByRef errMsg As String) As String
    Dim json As String
    json = RunQuery(baseUrl, lat, lon, outFields, whereClause, 0, errMsg)
    If Len(errMsg) > 0 Then Exit Function
    If FeatureCount(json) = 0 And Not HasArcgisError(json) Then
        json = RunQuery(baseUrl, lat, lon, outFields, whereClause, fallbackDistanceFt, errMsg)
    End If
    QueryWithFallback = json
End Function

' Read the search-buffer radius (in feet) from Setup, with a sane default
' and clamp. The cell can be anything (blank, text, a number out of range);
' fall back to DEFAULT_BUFFER_FEET unless we get a positive Long between
' 1 and 1000.
Public Function BufferFeet() As Long
    Dim raw As String, v As Double
    raw = SetupValue(NR_BUFFER)
    If Len(raw) = 0 Or Not IsNumeric(raw) Then
        BufferFeet = DEFAULT_BUFFER_FEET
        Exit Function
    End If
    v = CDbl(raw)
    If v < 1 Or v > 1000 Then
        BufferFeet = DEFAULT_BUFFER_FEET
        Exit Function
    End If
    BufferFeet = CLng(v)
End Function

Private Function AcubBufferFeet() As Long
    If BufferFeet() > ACUB_MIN_BUFFER_FEET Then
        AcubBufferFeet = BufferFeet()
    Else
        AcubBufferFeet = ACUB_MIN_BUFFER_FEET
    End If
End Function

Private Function RunQuery(ByVal baseUrl As String, ByVal lat As String, ByVal lon As String, _
        ByVal outFields As String, ByVal whereClause As String, ByVal distanceFt As Long, _
        ByRef errMsg As String) As String
    Dim url As String
    url = baseUrl & "/query?where=" & UrlEncode(whereClause) & _
        "&geometry=" & lon & "," & lat & _
        "&geometryType=esriGeometryPoint&inSR=4326" & _
        "&spatialRel=esriSpatialRelIntersects" & _
        "&outFields=" & UrlEncode(outFields) & _
        "&returnGeometry=false&f=json"
    If distanceFt > 0 Then url = url & "&distance=" & distanceFt & "&units=esriSRUnit_Foot"
    RunQuery = HttpGetText(url, errMsg)
End Function

Private Function BuildRouteName(ByVal routeJson As String) As String
    Dim desig As String, num As String, s As String
    desig = FirstString(routeJson, "RouteDesignation")
    num = FirstString(routeJson, "RouteNumber")
    s = Trim$(desig & " " & num)
    BuildRouteName = s
End Function

Private Function ClassLabels(ByVal codes As Collection) As String
    Dim seen As String, c As Variant, label As String, out As String
    If codes.Count = 0 Then
        ClassLabels = "No road segment within " & BufferFeet() & " ft"
        Exit Function
    End If
    For Each c In codes
        label = FunctionalSystemLabel(CLng(c))
        If InStr(seen, "|" & label & "|") = 0 Then
            seen = seen & "|" & label & "|"
            out = out & IIf(Len(out) > 0, " | ", "") & label
        End If
    Next c
    ClassLabels = out
End Function

' Federal Aid Status verdict per the §4.2 table. The label classifies the
' road, NOT the project — "Federal aid" means the road IS on the federal-
' aid system (Urban Minor Collector or higher), "Non-federal aid" means
' it isn't. The inspector decides what that means for their work order;
' we don't pre-judge eligibility.
Private Function FederalAidVerdict(ByVal codes As Collection, ByVal isUrban As Boolean) As String
    Dim c As Variant, isFederalAid As Boolean, manual As Boolean, hasRuralMinorCollector As Boolean
    Dim worstCode As Long: worstCode = 99

    If codes.Count = 0 Then
        FederalAidVerdict = "No road segment within " & BufferFeet() & " ft - review manually"
        Exit Function
    End If

    For Each c In codes
        Select Case CLng(c)
            Case 7   ' Local - never federal aid
            Case 6
                If isUrban Then
                    isFederalAid = True
                    If 6 < worstCode Then worstCode = 6
                Else
                    hasRuralMinorCollector = True   ' non-federal, but not "Local" either
                End If
            Case 1, 2, 3, 4, 5
                isFederalAid = True
                If CLng(c) < worstCode Then worstCode = CLng(c)
            Case Else: manual = True
        End Select
    Next c

    If isFederalAid Then
        FederalAidVerdict = "Federal aid - " & PrefixedClass(worstCode, isUrban)
    ElseIf manual Then
        FederalAidVerdict = "Review - non-certified class, check manually"
    ElseIf hasRuralMinorCollector Then
        ' Reachable only when every code present is 6 (rural) and/or 7 - if a
        ' 6 were urban it would have set isFederalAid above instead. Without
        ' this branch a Rural Minor Collector displayed as "...Rural Local",
        ' contradicting the FHWA Class column right next to it.
        FederalAidVerdict = "Non-federal aid - Rural Minor Collector"
    Else
        FederalAidVerdict = "Non-federal aid - " & IIf(isUrban, "Urban", "Rural") & " Local"
    End If
End Function

' Prefixes a federal-aid FHWA class (1-6; code 7/Local never reaches here)
' with Urban/Rural, matching the "Federal aid - <Urban/Rural class>" format
' in CLAUDE.md §4.2's eligibility table for every class, not just 4-6.
Private Function PrefixedClass(ByVal code As Long, ByVal isUrban As Boolean) As String
    PrefixedClass = IIf(isUrban, "Urban ", "Rural ") & FunctionalSystemLabel(code)
End Function
