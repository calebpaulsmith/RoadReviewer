Attribute VB_Name = "modClassify"
Option Explicit

' RoadReviewer V1 - Workflow 1: Classify Roads (§4.2 query strategy, F7 rule).
' For each Sites row: query the state's NFC layer (class + road name) and the
' nationwide NTAD ACUB layer (urban/rural), then write the eligibility verdict
' back to the row. All six Region V states are wired (MI/IN/WI since V1;
' MN/IL/OH added in PR #36, §4.2c-e); any other typed state code still gets
' the ACUB-only check (F8).

Private Const ACUB_OUTFIELDS As String = "NAME,UACE,state_1"

' The ACUB (urban-boundary) fallback search always uses at least this many
' feet, even if JobBufferFeet has been narrowed below it for precise road
' matching in a dense urban grid. Without this floor, narrowing the shared
' buffer for road-matching purposes would silently widen the "point right on
' the edge of the urban boundary gets flagged Rural" bug this exists to
' prevent - a site sitting a few feet on the wrong side of a road that
' itself touches the boundary should still resolve Urban.
Private Const ACUB_MIN_BUFFER_FEET As Long = 250

' Distance (ft) under which the two closest roads count as "close together" -
' close enough that which one the GPS point actually sits on is ambiguous, so
' a federal-aid second road there is worth surfacing (request 3). Module-level
' declarations MUST precede every Sub/Function (VBA "Only comments may appear
' after End Sub" rule, §9.3), so it lives up here, not next to its user.
Private Const CLOSE_ROAD_FEET As Double = 30

' Tracks whether the not-authoritative disclaimer dialog has been shown yet
' this Excel session. Module-level state must precede every procedure (VBA
' "Only comments may appear after End Sub" rule, §9.3).
Private mDisclaimerShown As Boolean

' The one primary action (friction fix: the old separate Geocode ->
' Classify ordering trap is gone). Check Roads geocodes any row that has
' an Address but no coordinates, then classifies every row.
Public Sub CheckRoads()
    ShowDisclaimerOnce
    ClassifyRows False
End Sub

' The inspector product carries no on-sheet disclaimer box (it was removed to
' keep Start Here focused); instead the same not-authoritative text is shown
' as a dialog the first time Check Roads runs each session. The standard
' product keeps its on-sheet box, so it doesn't need the dialog.
Private Sub ShowDisclaimerOnce()
    If gHeadless Then Exit Sub
    If mDisclaimerShown Then Exit Sub
    If Not ProductIsInspector() Then Exit Sub
    mDisclaimerShown = True
    MsgBox DisclaimerHeaderText() & vbCrLf & vbCrLf & DisclaimerBodyText(), _
        vbInformation, "Before you rely on these results"
End Sub

' Re-runs only rows whose Federal Aid Status starts with "Failed - ",
' which includes geocode failures (F12).
Public Sub ReRunFailedRows()
    ClassifyRows True
End Sub

Private Sub ClassifyRows(ByVal onlyFailed As Boolean)
    Dim ws As Worksheet, last As Long, r As Long
    Dim stateCode As String, stateHasNfc As Boolean
    Dim processed As Long, total As Long

    Set ws = SitesSheet()
    last = SitesLastRow()
    If last < SITES_FIRST_DATA_ROW Then
        If Not gHeadless Then MsgBox "No site rows found. Add points on the Sites sheet first.", vbInformation, "Check Roads"
        Exit Sub
    End If

    ' Blank State no longer silently means Michigan (PR #36 - with all six
    ' states wired there is no safe default). Mark every target row so the
    ' miss is visible + re-runnable, and tell the user where the State box is.
    stateCode = BareStateCode(SetupValue(NR_STATE))
    If Len(stateCode) = 0 Then
        For r = SITES_FIRST_DATA_ROW To last
            If Not RowIsEmpty(ws, r) Then
                If (Not onlyFailed) Or RowIsFailed(ws, r) Then _
                    ws.Cells(r, COL_ELIGIBILITY).Value = STATUS_FAILED_PREFIX & "no State selected"
            End If
        Next r
        ShowReviewerColumns
        If Not gHeadless Then MsgBox "Pick a State first (the State box on " & _
            IIf(ProductIsInspector(), "the Map Pages tab", "Start Here") & _
            "), then click Check Roads again.", vbExclamation, "Check Roads"
        Exit Sub
    End If
    ' Local var deliberately NOT named NfcWired - VBA identifiers are
    ' case-insensitive, so a local variable with the same name as the
    ' module-level NfcWired() function shadows it, and `NfcWired(stateCode)`
    ' on the right-hand side gets parsed as indexing a Boolean scalar like an
    ' array ("Expected array" compile error) instead of calling the function.
    stateHasNfc = NfcWired(stateCode)
    If Not stateHasNfc Then
        If Not gHeadless Then MsgBox "Road-class (NFC) lookup is not yet wired for " & stateCode & "." & vbCrLf & _
            "The ACUB urban-boundary check will still run on every row.", vbInformation, "Check Roads"
    End If

    For r = SITES_FIRST_DATA_ROW To last
        If Not RowIsEmpty(ws, r) Then
            If (Not onlyFailed) Or RowIsFailed(ws, r) Then total = total + 1
        End If
    Next r
    If total = 0 Then
        If Not gHeadless Then MsgBox IIf(onlyFailed, "No failed rows to re-run.", "No site rows to classify."), vbInformation, "Check Roads"
        Exit Sub
    End If

    ' Reveal the auto-check output columns BEFORE the loop starts (PR #37,
    ' user direction) so the user watches the verdicts appear as each row is
    ' classified, instead of the columns popping into view only at the end.
    ShowReviewerColumns
    ' Leave ScreenUpdating on and stay on the Sites sheet (not the Classify
    ' Roads control sheet the button lives on) so the inspector watches each
    ' row's cells fill in as it's classified - each row is dominated by
    ' network latency (2-3 HTTP calls), so redraw cost here is negligible.
    ws.Activate
    On Error GoTo Done
    For r = SITES_FIRST_DATA_ROW To last
        If RowIsEmpty(ws, r) Then GoTo NextRow
        If onlyFailed And Not RowIsFailed(ws, r) Then GoTo NextRow
        processed = processed + 1
        SetStatus "Classifying " & processed & " of " & total & " - " & _
            CStr(ws.Cells(r, COL_SITENAME).Value)
        ws.Cells(r, COL_SITENAME).Select
        ClassifyOneRow ws, r, stateCode
        DoEvents
NextRow:
    Next r

Done:
    ClearStatus
    If Err.Number <> 0 Then
        If Not gHeadless Then MsgBox "Classification stopped: " & Err.Description, vbExclamation, "Check Roads"
    Else
        If Not gHeadless Then MsgBox "Classified " & processed & " row(s).", vbInformation, "Check Roads"
    End If
End Sub

Private Function RowIsFailed(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    RowIsFailed = (InStr(1, CStr(ws.Cells(r, COL_ELIGIBILITY).Value), STATUS_FAILED_PREFIX, vbTextCompare) = 1)
End Function

' States whose NFC (road class) layer is wired in V1. Every state still gets
' the ACUB urban-boundary check regardless (F8).
Private Function NfcWired(ByVal stateCode As String) As Boolean
    Select Case stateCode
        Case "MI", "IN", "WI", "MN", "IL", "OH": NfcWired = True
        Case Else: NfcWired = False
    End Select
End Function

Private Sub ClassifyOneRow(ByVal ws As Worksheet, ByVal r As Long, ByVal stateCode As String)
    TraceLine "ClassifyOneRow row=" & r & " name=" & CStr(ws.Cells(r, COL_SITENAME).Value)
    ' Idempotent: clear prior lookup output before writing (N5).
    ClearLookupCells ws, r

    ' Auto-geocode (F4, folded into Check Roads): a row with an Address but
    ' no valid coordinates gets one Census geocoder pass first. Coordinates
    ' already typed into the row are never overwritten. A geocode failure
    ' lands in Federal Aid Status with the Failed prefix so Re-run Failed
    ' Rows retries it along with everything else.
    If Not HasValidCoords(ws, r) Then
        If Not IsBlank(ws.Cells(r, COL_ADDRESS).Value) Then
            Dim geoErr As String
            If Not GeocodeRow(ws, r, geoErr) Then
                ws.Cells(r, COL_ELIGIBILITY).Value = STATUS_FAILED_PREFIX & "geocode: " & geoErr
                TraceLine "  geocode failed: " & geoErr
                Exit Sub
            End If
        End If
    End If

    If Not HasValidCoords(ws, r) Then
        ws.Cells(r, COL_ELIGIBILITY).Value = STATUS_FAILED_PREFIX & "no/invalid coordinates"
        TraceLine "  no/invalid coords"
        Exit Sub
    End If

    Dim lat As String, lon As String, errMsg As String
    Dim latP As Double, lonP As Double
    lat = InvariantNum(ws.Cells(r, COL_LAT).Value)
    lon = InvariantNum(ws.Cells(r, COL_LON).Value)
    latP = CDbl(lat): lonP = CDbl(lon)

    ' --- ACUB (urban/rural), the single source of truth for Urban/Rural on
    ' every state (§4.2). Distinguish THREE outcomes: the point is inside an
    ' ACUB polygon (exactUrban, definitely Urban), it is outside but within
    ' the boundary buffer (boundaryAmbiguous - right on the urban edge), or
    ' clearly rural. The ambiguity feeds the yellow "Urban boundary edge"
    ' flag for Minor Collectors, whose federal-aid status flips on urban vs
    ' rural (request 2). ---
    Dim exactUrban As Boolean, boundaryAmbiguous As Boolean, acubName As String
    DetermineAcub lat, lon, exactUrban, boundaryAmbiguous, acubName, errMsg
    If Len(errMsg) > 0 Then
        ws.Cells(r, COL_ELIGIBILITY).Value = STATUS_FAILED_PREFIX & "ACUB query (" & errMsg & ")"
        Exit Sub
    End If
    ws.Cells(r, COL_URBANRURAL).Value = IIf(exactUrban, "Urban", "Rural")
    ws.Cells(r, COL_ACUBNAME).Value = acubName

    If Not NfcWired(stateCode) Then
        ws.Cells(r, COL_ELIGIBILITY).Value = "ACUB only - class lookup not wired for this state"
        Exit Sub
    End If

    ' --- Road segments (class + distance) + named roads (name + distance),
    ' all with geometry so we know how far each road is from the point. segs
    ' drives the federal-aid verdict; roads is the Road Name display list. ---
    Dim segs As Collection, roads As Collection
    Set roads = New Collection
    QueryStateRoads stateCode, lat, lon, latP, lonP, segs, roads, errMsg
    If Len(errMsg) > 0 Then
        ws.Cells(r, COL_ELIGIBILITY).Value = STATUS_FAILED_PREFIX & "NFC query (" & errMsg & ")"
        Exit Sub
    End If

    ' Census TIGER street names (with distances) - covers the local streets
    ' the state layers name poorly. Non-fatal; also fills the Street Name
    ' column. Adds its named roads into the merged Road Name list.
    Dim tigerErr As String, streetNames As String
    streetNames = AppendTigerRoads(lat, lon, latP, lonP, roads, tigerErr)
    ws.Cells(r, COL_STREET).Value = streetNames

    ' --- Displays + verdict ---
    ws.Cells(r, COL_ROADNAME).Value = FormatRoadList(roads)
    ws.Cells(r, COL_CLASS).Value = ClassLabelsFromSegs(segs)

    Dim verdict As String, reason As String
    ' Whether ANY named road was found near the point (state layer or TIGER).
    ' Lets ComputeVerdict distinguish "a road is here but the state assigns it
    ' no functional class" (unclassified - flag for review) from "genuinely
    ' nothing here" (request 4: Mohawk Trail is present in TIGER but absent
    ' from INDOT's class layer).
    ComputeVerdict segs, exactUrban, boundaryAmbiguous, (roads.Count > 0), verdict, reason
    ws.Cells(r, COL_ELIGIBILITY).Value = verdict
    ws.Cells(r, COL_REVIEWNOTE).Value = reason
End Sub

' Point-in-ACUB-polygon (exact) first; only if the point is outside any
' polygon do we buffer out to the boundary floor to detect "on the edge".
Private Sub DetermineAcub(ByVal lat As String, ByVal lon As String, _
        ByRef exactUrban As Boolean, ByRef boundaryAmbiguous As Boolean, _
        ByRef acubName As String, ByRef errMsg As String)
    Dim j As String
    j = RunQuery(ServiceUrl("ACUB"), lat, lon, ACUB_OUTFIELDS, "1=1", 0, errMsg, False)
    If Len(errMsg) > 0 Then Exit Sub
    If FeatureCount(j) > 0 Then
        exactUrban = True: boundaryAmbiguous = False: acubName = FirstString(j, "NAME")
        Exit Sub
    End If
    j = RunQuery(ServiceUrl("ACUB"), lat, lon, ACUB_OUTFIELDS, "1=1", AcubBufferFeet(), errMsg, False)
    If Len(errMsg) > 0 Then Exit Sub
    If FeatureCount(j) > 0 Then
        exactUrban = False: boundaryAmbiguous = True: acubName = FirstString(j, "NAME")
    Else
        exactUrban = False: boundaryAmbiguous = False: acubName = ""
    End If
End Sub

' Dispatch per wired state. Fills segs (each item Array(fhwaClass, distFt))
' and appends named roads (each item Array(name, distFt)) to roads.
Private Sub QueryStateRoads(ByVal stateCode As String, ByVal lat As String, ByVal lon As String, _
        ByVal latP As Double, ByVal lonP As Double, _
        ByRef segs As Collection, ByRef roads As Collection, ByRef errMsg As String)
    Set segs = New Collection
    Select Case stateCode
        Case "MI"
            ' Class from layer 353; trunkline route names from 543.
            AddClassSegs ServiceUrl("MI_NFC"), "FunctionalSystem", "RHRetireDate IS NULL", False, False, _
                lat, lon, latP, lonP, segs, errMsg
            If Len(errMsg) > 0 Then Exit Sub
            AddNamedRoads ServiceUrl("MI_ROUTE"), "RouteDesignation,RouteNumber", "RHRetireDate IS NULL", _
                "MI_ROUTE", lat, lon, latP, lonP, roads
        Case "IN"
            ' Authoritative INDOT Roads_and_Highways layer (§4.2a, switched
            ' 2026-07-16): UPPERCASE FUNCTIONAL_CLASS, where=1=1 - no
            ' record-status filter (the layer holds only statuses {1,4,5,null},
            ' none retired, and its null-status segments carry real classes,
            ' e.g. Wolf Run Rd Major Collector). Road name still comes from the
            ' separate centerlines layer (lowercase st_full, unchanged).
            AddClassSegs ServiceUrl("IN_NFC"), "FUNCTIONAL_CLASS", "1=1", False, False, _
                lat, lon, latP, lonP, segs, errMsg
            If Len(errMsg) > 0 Then Exit Sub
            AddNamedRoads ServiceUrl("IN_ROADNAME"), "st_full", "1=1", "SINGLE:st_full", lat, lon, latP, lonP, roads
        Case "WI"
            ' Local Road Network FIRST (it carries local roads AND most
            ' collectors - the large majority of points), then the State Trunk
            ' Network only if the local layer has no usable class for the point
            ' (PR "WI layer swap"). State highways appear in the local layer as
            ' unclassified "stubs" (null/0 class); AddWiLocalRoads skips those
            ' and reports sawStub, which - like an empty local result - triggers
            ' the trunk fallback so a point on a state highway is still
            ' classified. One query per layer (class + name together) keeps the
            ' hit count down on the local layer WisDOT is sensitive about.
            Dim sawStub As Boolean
            AddWiLocalRoads lat, lon, latP, lonP, segs, roads, sawStub, errMsg
            If Len(errMsg) > 0 Then Exit Sub
            If segs.Count = 0 Or sawStub Then
                AddClassSegs ServiceUrl("WI_STATE_TRUNK"), "FED_FC_CD", "1=1", True, False, _
                    lat, lon, latP, lonP, segs, errMsg
                If Len(errMsg) > 0 Then Exit Sub
                AddNamedRoads ServiceUrl("WI_STATE_TRUNK"), "HWYTYPE,HWYNUM,HWYDIR", "1=1", "WI_TRUNK", _
                    lat, lon, latP, lonP, roads
            End If
        ' MN/IL/OH (PR #36, §4.2c-e): bare FHWA 1-7 code, no active/retired
        ' filter, same shape as Indiana. MN and IL carry no street-name field
        ' (Census TIGER backfills names); OH's ROUTE_TYPE+ROUTE_NBR give
        ' trunkline names ("US 23") on the same layer.
        Case "MN"
            AddClassSegs ServiceUrl("MN_NFC"), "FUNCTIONAL_CLASS", "1=1", False, False, _
                lat, lon, latP, lonP, segs, errMsg
        Case "IL"
            ' FC is a STRING field ("1".."7") - isStringClass, like WI's FED_FC_CD.
            AddClassSegs ServiceUrl("IL_NFC"), "FC", "1=1", True, False, _
                lat, lon, latP, lonP, segs, errMsg
        Case "OH"
            AddClassSegs ServiceUrl("OH_NFC"), "FUNCTION_CLASS_CD", "1=1", False, False, _
                lat, lon, latP, lonP, segs, errMsg
            If Len(errMsg) > 0 Then Exit Sub
            AddNamedRoads ServiceUrl("OH_NFC"), "ROUTE_TYPE,ROUTE_NBR", "1=1", "OH_ROUTE", _
                lat, lon, latP, lonP, roads
    End Select
End Sub

' Wisconsin Local Road Network: one query returns both the class code and the
' street name (they share the layer), so this halves the local-layer hit count
' vs. the generic AddClassSegs+AddNamedRoads pair AND lets a state-highway stub
' be skipped for class and name in the same pass. A stub is any feature whose
' FNCT_CLS_CTGY_TYCD is null or decodes to 0 (an unclassified state-highway
' segment that appears here but is only classified on the trunk layer); each
' one sets sawStub so QueryStateRoads consults the trunk layer.
Private Sub AddWiLocalRoads(ByVal lat As String, ByVal lon As String, _
        ByVal latP As Double, ByVal lonP As Double, _
        ByRef segs As Collection, ByRef roads As Collection, _
        ByRef sawStub As Boolean, ByRef errMsg As String)
    Dim json As String, blocks As Collection, b As Variant
    Dim ints As Collection, cls As Long, nm As String, d As Double
    json = RunQuery(ServiceUrl("WI_LOCAL_ROADS"), lat, lon, _
        "FNCT_CLS_CTGY_TYCD,ST_LABL_NM", "1=1", BufferFeet(), errMsg, True)
    If Len(errMsg) > 0 Then Exit Sub
    Set blocks = FeatureBlocks(json)
    For Each b In blocks
        Set ints = ExtractIntegers(CStr(b), "FNCT_CLS_CTGY_TYCD")
        d = MinDistanceFt(CStr(b), lonP, latP)
        If ints.Count = 0 Then
            sawStub = True                          ' null class = state-highway stub
        Else
            cls = WisconsinLocalCategoryToFhwa(CLng(ints(1)))
            If cls >= 1 Then
                segs.Add Array(cls, d)
                nm = Trim$(FirstString(CStr(b), "ST_LABL_NM"))
                If Len(nm) > 0 Then roads.Add Array(nm, d)
            Else
                sawStub = True                      ' code 0 / unrecognized = stub
            End If
        End If
    Next b
End Sub

' Query a functional-class layer with geometry and append each segment's
' (fhwaClass, distanceFt) to segs. isStringClass=True reads the class as a
' string field (WI FED_FC_CD); wiLocalDecode=True runs the value through
' WisconsinLocalCategoryToFhwa (WI local-roads category code).
Private Sub AddClassSegs(ByVal baseUrl As String, ByVal classField As String, ByVal whereClause As String, _
        ByVal isStringClass As Boolean, ByVal wiLocalDecode As Boolean, _
        ByVal lat As String, ByVal lon As String, ByVal latP As Double, ByVal lonP As Double, _
        ByRef segs As Collection, ByRef errMsg As String)
    Dim json As String, blocks As Collection, b As Variant, cls As Long
    json = RunQuery(baseUrl, lat, lon, classField, whereClause, BufferFeet(), errMsg, True)
    If Len(errMsg) > 0 Then Exit Sub
    Set blocks = FeatureBlocks(json)
    For Each b In blocks
        cls = ClassFromBlock(CStr(b), classField, isStringClass, wiLocalDecode)
        If cls >= 0 Then segs.Add Array(cls, MinDistanceFt(CStr(b), lonP, latP))
    Next b
End Sub

Private Function ClassFromBlock(ByVal block As String, ByVal classField As String, _
        ByVal isStringClass As Boolean, ByVal wiLocalDecode As Boolean) As Long
    Dim v As String, ints As Collection
    If isStringClass Then
        v = FirstString(block, classField)
    Else
        Set ints = ExtractIntegers(block, classField)
        If ints.Count > 0 Then v = CStr(ints(1))
    End If
    If Len(v) = 0 Or Not IsNumeric(v) Then ClassFromBlock = -1: Exit Function
    If wiLocalDecode Then
        ClassFromBlock = WisconsinLocalCategoryToFhwa(CLng(v))
    Else
        ClassFromBlock = CLng(v)
    End If
End Function

' Query a name-bearing layer with geometry and append each named road's
' (name, distanceFt) to roads. nameMode picks how the name is built from the
' feature block.
Private Sub AddNamedRoads(ByVal baseUrl As String, ByVal outFields As String, ByVal whereClause As String, _
        ByVal nameMode As String, ByVal lat As String, ByVal lon As String, _
        ByVal latP As Double, ByVal lonP As Double, ByRef roads As Collection)
    Dim json As String, blocks As Collection, b As Variant, nm As String, e As String
    json = RunQuery(baseUrl, lat, lon, outFields, whereClause, BufferFeet(), e, True)
    If Len(e) > 0 Then Exit Sub       ' road names are non-fatal
    Set blocks = FeatureBlocks(json)
    For Each b In blocks
        nm = RoadNameFromBlock(CStr(b), nameMode)
        If Len(nm) > 0 Then roads.Add Array(nm, MinDistanceFt(CStr(b), lonP, latP))
    Next b
End Sub

Private Function RoadNameFromBlock(ByVal block As String, ByVal nameMode As String) As String
    Dim rt As String, rn As String
    Select Case True
        Case nameMode = "MI_ROUTE"
            RoadNameFromBlock = Trim$(FirstString(block, "RouteDesignation") & " " & FirstString(block, "RouteNumber"))
        Case nameMode = "WI_TRUNK"
            RoadNameFromBlock = Trim$(FirstString(block, "HWYTYPE") & " " & FirstString(block, "HWYNUM") & " " & FirstString(block, "HWYDIR"))
        Case nameMode = "OH_ROUTE"
            ' Only the recognizable route systems: IR (interstate), US, SR.
            ' County/township/municipal codes ("MR 00923") read as gibberish
            ' to an inspector - TIGER supplies those street names instead.
            rt = UCase$(Trim$(FirstString(block, "ROUTE_TYPE")))
            rn = Trim$(FirstString(block, "ROUTE_NBR"))
            If (rt = "IR" Or rt = "US" Or rt = "SR") And IsNumeric(rn) Then
                RoadNameFromBlock = rt & " " & CStr(CLng(rn))   ' "US 00023" -> "US 23"
            End If
        Case Left$(nameMode, 7) = "SINGLE:"
            RoadNameFromBlock = Trim$(FirstString(block, Mid$(nameMode, 8)))
    End Select
End Function

' Census TIGER Local Roads (layer 8), with geometry. Appends each street's
' (name, distanceFt) to roads AND returns a plain pipe-joined name list for
' the Street Name column. Non-fatal on failure.
Private Function AppendTigerRoads(ByVal lat As String, ByVal lon As String, _
        ByVal latP As Double, ByVal lonP As Double, ByRef roads As Collection, _
        ByRef errMsg As String) As String
    Dim json As String, blocks As Collection, b As Variant, nm As String
    Dim seen As String, out As String
    json = RunQuery(ServiceUrl("TIGER_ROADS"), lat, lon, "NAME", "1=1", BufferFeet(), errMsg, True)
    If Len(errMsg) > 0 Then Exit Function
    Set blocks = FeatureBlocks(json)
    For Each b In blocks
        nm = Trim$(FirstString(CStr(b), "NAME"))
        If Len(nm) > 0 Then
            roads.Add Array(nm, MinDistanceFt(CStr(b), lonP, latP))
            If InStr(seen, "|" & LCase$(nm) & "|") = 0 Then
                seen = seen & "|" & LCase$(nm) & "|"
                out = out & IIf(Len(out) > 0, " | ", "") & nm
            End If
        End If
    Next b
    AppendTigerRoads = out
End Function

' Merge the collected roads into one "Name (D ft) | Name (D ft)" string:
' dedup by NORMALIZED name (so "N Meridian St", "MERIDIAN ST" and the same
' road from two sources collapse to one entry - request 3), keep the nearest
' distance, prefer a mixed-case display over ALL-CAPS, sort nearest-first.
Private Function FormatRoadList(ByVal roads As Collection) As String
    Dim dict As Object, rv As Variant, nm As String, d As Double, key As String
    Dim cur As Variant
    Set dict = CreateObject("Scripting.Dictionary")
    For Each rv In roads
        nm = Trim$(CStr(rv(0))): d = CDbl(rv(1))
        If Len(nm) > 0 Then
            key = NormalizeRoadKey(nm)
            If Len(key) = 0 Then key = LCase$(nm)
            If Not dict.Exists(key) Then
                ' item = Array(displayName, minDist, displayIsMixedCase)
                dict.Add key, Array(nm, d, IsMixedCase(nm))
            Else
                cur = dict(key)
                Dim newName As String, newDist As Double, newMixed As Boolean
                newName = CStr(cur(0)): newDist = CDbl(cur(1)): newMixed = CBool(cur(2))
                If d < newDist Then newDist = d
                ' Prefer a mixed-case name; among same case-class, prefer nearer.
                If IsMixedCase(nm) And Not newMixed Then
                    newName = nm: newMixed = True
                ElseIf (IsMixedCase(nm) = newMixed) And d < CDbl(cur(1)) Then
                    newName = nm
                End If
                dict(key) = Array(newName, newDist, newMixed)
            End If
        End If
    Next rv
    If dict.Count = 0 Then Exit Function

    ' Collect + insertion-sort by distance.
    Dim arr() As Variant, cnt As Long, k As Variant, i As Long, j As Long, tmp As Variant
    ReDim arr(0 To dict.Count - 1)
    cnt = 0
    For Each k In dict.Keys
        arr(cnt) = dict(k): cnt = cnt + 1
    Next k
    For i = 1 To cnt - 1
        tmp = arr(i): j = i - 1
        Do While j >= 0
            If CDbl(arr(j)(1)) <= CDbl(tmp(1)) Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i

    Dim out As String
    For i = 0 To cnt - 1
        out = out & IIf(Len(out) > 0, " | ", "") & arr(i)(0) & " (" & Format$(arr(i)(1), "0") & " ft)"
    Next i
    FormatRoadList = out
End Function

' True if the name carries at least one lowercase letter (i.e. is not
' ALL-CAPS). Used to prefer TIGER-style "Harrison Pkwy" over the state
' centerline's "HARRISON PKY" when both name the same road.
Private Function IsMixedCase(ByVal s As String) As Boolean
    IsMixedCase = (StrComp(s, UCase$(s), vbBinaryCompare) <> 0)
End Function

' Canonical dedup key for a road name so the same physical road written two
' ways collapses to one Road Name entry (request 3, test 7.16). Steps:
'   1. upper-case, drop punctuation, collapse whitespace;
'   2. drop a single leading AND trailing compass token (N/S/E/W/NE/.., or the
'      spelled-out word) - "N Meridian St" and "Meridian St" are one road, and
'      "E Market St"/"W Market St" collapse to the same key too;
'   3. canonicalize the street-type suffix (ST/STREET, RD/ROAD, PKWY/PKY, ..)
'      so abbreviation differences don't split a road.
Private Function NormalizeRoadKey(ByVal name As String) As String
    Dim s As String, parts() As String, toks As Collection, i As Long, t As String
    s = UCase$(Trim$(name))
    ' Strip punctuation to spaces.
    Dim ch As String, clean As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            clean = clean & ch
        Else
            clean = clean & " "
        End If
    Next i
    ' Tokenize on runs of spaces.
    Set toks = New Collection
    parts = Split(clean, " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then toks.Add parts(i)
    Next i
    If toks.Count = 0 Then Exit Function
    ' Drop a leading compass token (keep at least one non-compass token).
    If toks.Count > 1 Then
        If IsCompassToken(CStr(toks(1))) Then toks.Remove 1
    End If
    ' Drop a trailing compass token.
    If toks.Count > 1 Then
        If IsCompassToken(CStr(toks(toks.Count))) Then toks.Remove toks.Count
    End If
    ' Canonicalize the final (street-type) token.
    If toks.Count > 1 Then
        t = CanonStreetType(CStr(toks(toks.Count)))
        toks.Remove toks.Count
        toks.Add t
    End If
    Dim out As String
    For i = 1 To toks.Count
        out = out & IIf(Len(out) > 0, " ", "") & toks(i)
    Next i
    NormalizeRoadKey = out
End Function

Private Function IsCompassToken(ByVal t As String) As Boolean
    Select Case UCase$(t)
        Case "N", "S", "E", "W", "NE", "NW", "SE", "SW", _
             "NORTH", "SOUTH", "EAST", "WEST": IsCompassToken = True
        Case Else: IsCompassToken = False
    End Select
End Function

' Map common street-type synonyms/abbreviations to one canonical token so
' "PKWY"/"PKY"/"PARKWAY" (etc.) don't split a road across the dedup.
Private Function CanonStreetType(ByVal t As String) As String
    Select Case UCase$(t)
        Case "ST", "STREET": CanonStreetType = "ST"
        Case "RD", "ROAD": CanonStreetType = "RD"
        Case "AVE", "AV", "AVENUE": CanonStreetType = "AVE"
        Case "PKWY", "PKY", "PARKWAY", "PKWAY": CanonStreetType = "PKWY"
        Case "DR", "DRIVE": CanonStreetType = "DR"
        Case "LN", "LANE": CanonStreetType = "LN"
        Case "CT", "COURT": CanonStreetType = "CT"
        Case "BLVD", "BOULEVARD": CanonStreetType = "BLVD"
        Case "TRL", "TRAIL": CanonStreetType = "TRL"
        Case "HWY", "HIGHWAY": CanonStreetType = "HWY"
        Case "CIR", "CIRCLE": CanonStreetType = "CIR"
        Case "PL", "PLACE": CanonStreetType = "PL"
        Case "TER", "TERR", "TERRACE": CanonStreetType = "TER"
        Case "PT", "POINT": CanonStreetType = "PT"
        Case "SQ", "SQUARE": CanonStreetType = "SQ"
        Case "CV", "COVE": CanonStreetType = "CV"
        Case "WAY": CanonStreetType = "WAY"
        Case Else: CanonStreetType = UCase$(t)
    End Select
End Function

' Distinct FHWA class labels across the detected segments, nearest-first.
Private Function ClassLabelsFromSegs(ByVal segs As Collection) As String
    If segs.Count = 0 Then
        ClassLabelsFromSegs = "No road segment within " & BufferFeet() & " ft"
        Exit Function
    End If
    Dim sorted() As Variant, seen As String, out As String, i As Long, label As String
    sorted = SortSegsByDist(segs)
    For i = 0 To UBound(sorted)
        label = FunctionalSystemLabel(CLng(sorted(i)(0)))
        If InStr(seen, "|" & label & "|") = 0 Then
            seen = seen & "|" & label & "|"
            out = out & IIf(Len(out) > 0, " | ", "") & label
        End If
    Next i
    ClassLabelsFromSegs = out
End Function

' The core verdict (requests 1-3). Yellow only ever DOWNGRADES green ("looks
' non-federal but an ambiguity could make it federal - review"); a road whose
' own closest segment is federal-aid stays RED. Reason (<=3 words) goes to the
' Review Reason column and is echoed in the "Review - ..." status so the row
' tints yellow.
Private Sub ComputeVerdict(ByVal segs As Collection, ByVal exactUrban As Boolean, _
        ByVal boundaryAmbiguous As Boolean, ByVal namedRoadNearby As Boolean, _
        ByRef verdict As String, ByRef reason As String)
    reason = ""
    If segs.Count = 0 Then
        ' A named road is present nearby but the state layer assigns it no
        ' functional class (request 4). Surface it for manual review rather
        ' than reporting "no road" - the road IS there (see the Road Name
        ' column), it's just unclassified in the state's data.
        If namedRoadNearby Then
            verdict = "Review - road not classified"
            reason = "Unclassified road"
        Else
            verdict = "Review - no road within " & BufferFeet() & " ft"
            reason = "No road found"
        End If
        Exit Sub
    End If

    Dim sorted() As Variant
    sorted = SortSegsByDist(segs)          ' nearest-first
    Dim pClass As Long, pDist As Double
    pClass = CLng(sorted(0)(0)): pDist = CDbl(sorted(0)(1))

    ' The road the point is ON (closest segment) drives red/green.
    If pClass = 0 Then
        verdict = "Review - non-certified class, check manually"
        reason = "Non-certified"
        Exit Sub
    End If
    If ClassIsFederal(pClass, exactUrban) Then
        verdict = "Federal aid - " & PrefixedClass(pClass, exactUrban)   ' RED stays RED
        Exit Sub
    End If

    ' Primary is non-federal -> green base. Build its label.
    Dim baseText As String
    If pClass = 6 Then
        baseText = "Non-federal aid - Rural Minor Collector"   ' non-fed 6 => rural
    ElseIf pClass = 7 Then
        baseText = "Non-federal aid - " & IIf(exactUrban, "Urban", "Rural") & " Local"
    Else
        baseText = "Non-federal aid - " & PrefixedClass(pClass, exactUrban)
    End If

    ' Ambiguities that turn green -> yellow, most-specific reason first.
    ' (a) the 2nd-closest road is within 30 ft of the closest AND is federal.
    If UBound(sorted) >= 1 Then
        If (CDbl(sorted(1)(1)) - pDist) < CLOSE_ROAD_FEET And ClassIsFederal(CLng(sorted(1)(0)), exactUrban) Then
            reason = "Second road close"
        End If
    End If
    ' (b) any other detected road is federal-aid.
    If Len(reason) = 0 Then
        Dim i As Long
        For i = 1 To UBound(sorted)
            If ClassIsFederal(CLng(sorted(i)(0)), exactUrban) Then reason = "Nearby FHWA road": Exit For
        Next i
    End If
    ' (c) a Minor Collector sitting on the urban boundary (urban would flip it
    ' to federal). Only matters for class 6.
    If Len(reason) = 0 And pClass = 6 And boundaryAmbiguous Then
        reason = "Urban boundary edge"
    End If

    If Len(reason) > 0 Then
        verdict = "Review - " & reason
    Else
        verdict = baseText
    End If
End Sub

' A segment's class is federal-aid: 1-5 always; 6 (Minor Collector) only when
' urban; 7 (Local) never; 0 (non-certified) is handled as "review" upstream.
Private Function ClassIsFederal(ByVal cls As Long, ByVal isUrban As Boolean) As Boolean
    Select Case cls
        Case 1, 2, 3, 4, 5: ClassIsFederal = True
        Case 6: ClassIsFederal = isUrban
        Case Else: ClassIsFederal = False
    End Select
End Function

' Insertion-sort a segs Collection (items Array(class, dist)) into a
' distance-ascending Variant array.
Private Function SortSegsByDist(ByVal segs As Collection) As Variant()
    Dim arr() As Variant, i As Long, j As Long, tmp As Variant, s As Variant
    ReDim arr(0 To segs.Count - 1)
    i = 0
    For Each s In segs
        arr(i) = s: i = i + 1
    Next s
    For i = 1 To UBound(arr)
        tmp = arr(i): j = i - 1
        Do While j >= 0
            If CDbl(arr(j)(1)) <= CDbl(tmp(1)) Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
    SortSegsByDist = arr
End Function

Private Sub ClearLookupCells(ByVal ws As Worksheet, ByVal r As Long)
    ws.Cells(r, COL_CLASS).ClearContents
    ws.Cells(r, COL_STREET).ClearContents
    ws.Cells(r, COL_URBANRURAL).ClearContents
    ws.Cells(r, COL_ACUBNAME).ClearContents
    ws.Cells(r, COL_ROADNAME).ClearContents
    ws.Cells(r, COL_ELIGIBILITY).ClearContents
    ws.Cells(r, COL_REVIEWNOTE).ClearContents
End Sub

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

' One ArcGIS query. withGeom adds returnGeometry=true&outSR=4326 so the caller
' can compute per-feature distances; otherwise geometry is omitted (ACUB
' point-in-polygon only needs the count + NAME).
Private Function RunQuery(ByVal baseUrl As String, ByVal lat As String, ByVal lon As String, _
        ByVal outFields As String, ByVal whereClause As String, ByVal distanceFt As Long, _
        ByRef errMsg As String, Optional ByVal withGeom As Boolean = False) As String
    Dim url As String, geomPart As String
    If withGeom Then geomPart = "&returnGeometry=true&outSR=4326" Else geomPart = "&returnGeometry=false"
    url = baseUrl & "/query?where=" & UrlEncode(whereClause) & _
        "&geometry=" & lon & "," & lat & _
        "&geometryType=esriGeometryPoint&inSR=4326" & _
        "&spatialRel=esriSpatialRelIntersects" & _
        "&outFields=" & UrlEncode(outFields) & geomPart & "&f=json"
    If distanceFt > 0 Then url = url & "&distance=" & distanceFt & "&units=esriSRUnit_Foot"
    RunQuery = HttpGetText(url, errMsg)
End Function

' Prefixes an FHWA class label with Urban/Rural, matching the
' "Federal aid - <Urban/Rural class>" format in CLAUDE.md §4.2's table.
Private Function PrefixedClass(ByVal code As Long, ByVal isUrban As Boolean) As String
    PrefixedClass = IIf(isUrban, "Urban ", "Rural ") & FunctionalSystemLabel(code)
End Function
