Attribute VB_Name = "modClassify"
Option Explicit

' RoadReviewer V1 - Workflow 1: Classify Roads (§4.2 query strategy, F7 rule).
' For each Sites row: query the MDOT NFC layer (class), the route layer (name)
' and the nationwide NTAD ACUB layer (urban/rural), then write the eligibility
' verdict back to the row. Michigan is wired in V1; other states still get the
' ACUB check (F8).

Private Const NFC_OUTFIELDS As String = "FunctionalSystem,PR"
Private Const ROUTE_OUTFIELDS As String = "RouteDesignation,RouteNumber"
Private Const ACUB_OUTFIELDS As String = "NAME,UACE,state_1"

Public Sub ClassifyAllRows()
    ClassifyRows False
End Sub

Public Sub ReRunFailedClassifications()
    ClassifyRows True
End Sub

Private Sub ClassifyRows(ByVal onlyFailed As Boolean)
    Dim ws As Worksheet, last As Long, r As Long
    Dim stateCode As String, nfcWired As Boolean
    Dim processed As Long, total As Long

    Set ws = SitesSheet()
    last = SitesLastRow()
    If last < SITES_FIRST_DATA_ROW Then
        MsgBox "No site rows found. Add points on the Sites sheet first.", vbInformation, "Classify Roads"
        Exit Sub
    End If

    stateCode = UCase$(SetupValue(NR_STATE))
    If Len(stateCode) = 0 Then stateCode = "MI"
    nfcWired = (stateCode = "MI")
    If Not nfcWired Then
        MsgBox "Road-class (NFC) lookup is not yet wired for " & stateCode & "." & vbCrLf & _
            "The ACUB urban-boundary check will still run on every row.", vbInformation, "Classify Roads"
    End If

    For r = SITES_FIRST_DATA_ROW To last
        If Not RowIsEmpty(ws, r) Then
            If (Not onlyFailed) Or RowIsFailed(ws, r) Then total = total + 1
        End If
    Next r
    If total = 0 Then
        MsgBox IIf(onlyFailed, "No failed rows to re-run.", "No site rows to classify."), vbInformation, "Classify Roads"
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
        ClassifyOneRow ws, r, nfcWired
        DoEvents
NextRow:
    Next r

Done:
    Application.ScreenUpdating = True
    ClearStatus
    If Err.Number <> 0 Then
        MsgBox "Classification stopped: " & Err.Description, vbExclamation, "Classify Roads"
    Else
        MsgBox "Classified " & processed & " row(s).", vbInformation, "Classify Roads"
    End If
End Sub

Private Function RowIsFailed(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    RowIsFailed = (InStr(1, CStr(ws.Cells(r, COL_ELIGIBILITY).Value), STATUS_FAILED_PREFIX, vbTextCompare) = 1)
End Function

Private Sub ClassifyOneRow(ByVal ws As Worksheet, ByVal r As Long, ByVal nfcWired As Boolean)
    ' Idempotent: clear prior lookup output before writing (N5).
    ClearLookupCells ws, r

    If Not HasValidCoords(ws, r) Then
        ws.Cells(r, COL_ELIGIBILITY).Value = STATUS_FAILED_PREFIX & "no/invalid coordinates"
        Exit Sub
    End If

    Dim lat As String, lon As String, errMsg As String
    lat = InvariantNum(ws.Cells(r, COL_LAT).Value)
    lon = InvariantNum(ws.Cells(r, COL_LON).Value)

    ' --- ACUB (urban/rural) - runs for every state ---
    Dim acubJson As String, isUrban As Boolean, acubName As String
    acubJson = QueryWithFallback(REST_ACUB, lat, lon, ACUB_OUTFIELDS, False, errMsg)
    If Len(errMsg) > 0 Then
        ws.Cells(r, COL_ELIGIBILITY).Value = STATUS_FAILED_PREFIX & "ACUB query (" & errMsg & ")"
        Exit Sub
    End If
    isUrban = (FeatureCount(acubJson) > 0)
    If isUrban Then acubName = FirstString(acubJson, "NAME")
    ws.Cells(r, COL_URBANRURAL).Value = IIf(isUrban, "Urban", "Rural")
    ws.Cells(r, COL_ACUBNAME).Value = acubName

    If Not nfcWired Then
        ws.Cells(r, COL_ELIGIBILITY).Value = "ACUB only - class lookup not wired for this state"
        Exit Sub
    End If

    ' --- NFC class (MDOT layer 353) ---
    Dim nfcJson As String, codes As Collection
    nfcJson = QueryWithFallback(REST_MDOT_NFC, lat, lon, NFC_OUTFIELDS, True, errMsg)
    If Len(errMsg) > 0 Then
        ws.Cells(r, COL_ELIGIBILITY).Value = STATUS_FAILED_PREFIX & "NFC query (" & errMsg & ")"
        Exit Sub
    End If
    Set codes = ExtractIntegers(nfcJson, "FunctionalSystem")

    ' --- Route name (MDOT layer 543; blank for local streets) ---
    Dim routeJson As String, roadName As String
    routeJson = QueryWithFallback(REST_MDOT_ROUTE, lat, lon, ROUTE_OUTFIELDS, True, errMsg)
    If Len(errMsg) = 0 Then roadName = BuildRouteName(routeJson)
    ws.Cells(r, COL_ROADNAME).Value = roadName

    ' --- Class label + eligibility verdict ---
    ws.Cells(r, COL_CLASS).Value = ClassLabels(codes)
    ws.Cells(r, COL_ELIGIBILITY).Value = EligibilityVerdict(codes, isUrban)
End Sub

Private Sub ClearLookupCells(ByVal ws As Worksheet, ByVal r As Long)
    ws.Cells(r, COL_CLASS).ClearContents
    ws.Cells(r, COL_URBANRURAL).ClearContents
    ws.Cells(r, COL_ACUBNAME).ClearContents
    ws.Cells(r, COL_ROADNAME).ClearContents
    ws.Cells(r, COL_ELIGIBILITY).ClearContents
End Sub

' Exact point intersect first; if no hit, retry with a 150-ft buffer (§4.2).
Private Function QueryWithFallback(ByVal baseUrl As String, ByVal lat As String, ByVal lon As String, _
        ByVal outFields As String, ByVal retiredFilter As Boolean, ByRef errMsg As String) As String
    Dim json As String
    json = RunQuery(baseUrl, lat, lon, outFields, retiredFilter, 0, errMsg)
    If Len(errMsg) > 0 Then Exit Function
    If FeatureCount(json) = 0 And Not HasArcgisError(json) Then
        json = RunQuery(baseUrl, lat, lon, outFields, retiredFilter, 150, errMsg)
    End If
    QueryWithFallback = json
End Function

Private Function RunQuery(ByVal baseUrl As String, ByVal lat As String, ByVal lon As String, _
        ByVal outFields As String, ByVal retiredFilter As Boolean, ByVal distanceFt As Long, _
        ByRef errMsg As String) As String
    Dim url As String, whereClause As String
    whereClause = IIf(retiredFilter, "RHRetireDate IS NULL", "1=1")
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
        ClassLabels = "No road segment within 150 ft"
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

' Eligibility per the §4.2 table. Ineligible if ANY segment is ineligible.
Private Function EligibilityVerdict(ByVal codes As Collection, ByVal isUrban As Boolean) As String
    Dim c As Variant, ineligible As Boolean, manual As Boolean
    Dim worstCode As Long: worstCode = 99

    If codes.Count = 0 Then
        EligibilityVerdict = "No road segment within 150 ft - manual review"
        Exit Function
    End If

    For Each c In codes
        Select Case CLng(c)
            Case 7:                 ' Local - always eligible
            Case 6: If isUrban Then ineligible = True: If 6 < worstCode Then worstCode = 6
            Case 1, 2, 3, 4, 5
                ineligible = True
                If CLng(c) < worstCode Then worstCode = CLng(c)
            Case Else: manual = True
        End Select
    Next c

    If ineligible Then
        EligibilityVerdict = "INELIGIBLE - " & PrefixedClass(worstCode, isUrban)
    ElseIf manual Then
        EligibilityVerdict = "REVIEW - non-certified class, check manually"
    Else
        EligibilityVerdict = "ELIGIBLE - " & IIf(isUrban, "Urban", "Rural") & " Local"
    End If
End Function

Private Function PrefixedClass(ByVal code As Long, ByVal isUrban As Boolean) As String
    Dim prefix As String
    prefix = IIf(isUrban, "Urban ", "Rural ")
    Select Case code
        Case 6: PrefixedClass = prefix & "Minor Collector"
        Case 5: PrefixedClass = prefix & "Major Collector"
        Case 4: PrefixedClass = prefix & "Minor Arterial"
        Case Else: PrefixedClass = FunctionalSystemLabel(code)
    End Select
End Function
