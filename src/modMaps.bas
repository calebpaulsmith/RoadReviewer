Attribute VB_Name = "modMaps"
Option Explicit

' RoadReviewer V1 - Workflow 3: Maps & FIRMettes (F3.3, F10).
' Ports the FIRMette download + Map Pages flow from the Site Inspector
' prototype but reads from the consolidated Sites table and the Setup
' named ranges instead of input boxes / a separate WO/DI column pair.
'
' Network calls all funnel through modHttp.HttpGetText and HttpDownloadPdf
' so MDOT-style 403s never bite us (browser UA) and every URL ends up in
' the trace file when gTracePath is set.

' ---- output folder (§8.9) -------------------------------------------------

Public Sub SelectOutputFolder()
    Dim fd As Object, chosen As String
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Choose the output folder for FIRMettes and map PDFs"
    If fd.Show = -1 Then
        chosen = fd.SelectedItems(1)
        If Right$(chosen, 1) <> "\" Then chosen = chosen & "\"
        ThisWorkbook.Names(NR_OUTFOLDER).RefersToRange.Value = chosen
    End If
End Sub

' Effective output folder: the Setup value if set, else the default pattern.
Public Function ResolveOutputFolder() As String
    Dim v As String
    v = SetupValue(NR_OUTFOLDER)
    If Len(v) > 0 Then
        If Right$(v, 1) <> "\" Then v = v & "\"
        ResolveOutputFolder = v
    Else
        ResolveOutputFolder = DefaultOutputFolder()
    End If
End Function

' {base}\Desktop\Script\RoadReviewer\{Disaster}\{WO-DI}\  (§8.9)
' Each segment is omitted when its corresponding Setup value is blank:
'   WO=123 DI=456 Disaster=DR-TEST -> ...\RoadReviewer\DR-TEST\WO123-DI456\
'   WO=123 DI=""  Disaster=DR-TEST -> ...\RoadReviewer\DR-TEST\WO123\
'   WO=""  DI=""  Disaster=""       -> ...\RoadReviewer\
Private Function DefaultOutputFolder() As String
    Dim profile As String, base As String, disaster As String, wo As String, di As String, jobSeg As String
    profile = Environ$("USERPROFILE")
    If FolderExists(profile & "\OneDrive - FEMA") Then
        base = profile & "\OneDrive - FEMA"
    ElseIf FolderExists(profile & "\OneDrive") Then
        base = profile & "\OneDrive"
    Else
        base = profile
    End If
    disaster = CleanFileName(SetupValue(NR_DISASTER))
    wo = CleanFileName(SetupValue(NR_WO))
    di = CleanFileName(SetupValue(NR_DI))
    jobSeg = JobIds(wo, di, "-", "WO", "DI")
    DefaultOutputFolder = base & "\Desktop\Script\RoadReviewer\" & _
        IIf(Len(disaster) > 0, disaster & "\", "") & _
        IIf(Len(jobSeg) > 0, jobSeg & "\", "")
End Function

Private Function FolderExists(ByVal path As String) As Boolean
    FolderExists = (Len(Dir$(path, vbDirectory)) > 0)
End Function

' Create the folder (and any missing parents) on demand.
Public Function EnsureFolderExists(ByVal path As String) As Boolean
    Dim fso As Object
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    CreateTree fso, path
    EnsureFolderExists = fso.FolderExists(path)
    Exit Function
Fail:
    EnsureFolderExists = False
End Function

Private Sub CreateTree(ByVal fso As Object, ByVal path As String)
    Dim parent As String
    path = TrimSlash(path)
    If Len(path) = 0 Then Exit Sub
    If fso.FolderExists(path) Then Exit Sub
    parent = fso.GetParentFolderName(path)
    If Len(parent) > 0 And Not fso.FolderExists(parent) Then CreateTree fso, parent
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Private Function TrimSlash(ByVal s As String) As String
    Do While Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    TrimSlash = s
End Function

' ---- FIRMette download (FEMA Print FIRMette GP service) -------------------

Public Sub DownloadFirmettes()
    FirmetteRunRows False
End Sub

Public Sub ReRunFailedFirmettes()
    FirmetteRunRows True
End Sub

' Drives the FIRMette download for every (or just-failed) Sites row.
Private Sub FirmetteRunRows(ByVal onlyFailed As Boolean)
    Dim ws As Worksheet, last As Long, r As Long
    Dim folder As String, disaster As String, wo As String, di As String
    Dim processed As Long, ok As Long, failed As Long, total As Long
    Dim msg As String, fileName As String, fullPath As String
    Dim lat As String, lon As String, siteName As String

    Set ws = SitesSheet()
    last = SitesLastRow()
    If last < SITES_FIRST_DATA_ROW Then
        If Not gHeadless Then MsgBox "No site rows found. Add points on the Sites sheet first.", vbInformation, "Download FIRMettes"
        Exit Sub
    End If

    folder = ResolveOutputFolder()
    If Not EnsureFolderExists(folder) Then
        If Not gHeadless Then MsgBox "Could not create the output folder:" & vbCrLf & folder, vbExclamation, "Download FIRMettes"
        Exit Sub
    End If
    disaster = SetupValue(NR_DISASTER)
    wo = SetupValue(NR_WO)
    di = SetupValue(NR_DI)

    For r = SITES_FIRST_DATA_ROW To last
        If ShouldRunFirmetteRow(ws, r, onlyFailed) Then total = total + 1
    Next r
    If total = 0 Then
        If Not gHeadless Then MsgBox IIf(onlyFailed, "No failed FIRMette rows to re-run.", "No site rows to process."), vbInformation, "Download FIRMettes"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    For r = SITES_FIRST_DATA_ROW To last
        If Not ShouldRunFirmetteRow(ws, r, onlyFailed) Then GoTo NextRow
        processed = processed + 1
        siteName = CStr(ws.Cells(r, COL_SITENAME).Value)
        If Len(Trim$(siteName)) = 0 Then siteName = "row" & r
        SetStatus "Downloading FIRMette " & processed & " of " & total & " - " & siteName
        DoEvents

        lat = InvariantNum(ws.Cells(r, COL_LAT).Value)
        lon = InvariantNum(ws.Cells(r, COL_LON).Value)
        fileName = FirmetteFileName(wo, di, disaster, siteName)
        fullPath = folder & fileName

        msg = ""
        If DownloadOneFirmette(lat, lon, fullPath, msg) Then
            ws.Cells(r, COL_FIRMSTATUS).Value = "Downloaded: " & fileName
            ok = ok + 1
        Else
            ws.Cells(r, COL_FIRMSTATUS).Value = STATUS_FAILED_PREFIX & Left$(msg, 240)
            failed = failed + 1
        End If
        DoEvents
NextRow:
    Next r
    Application.ScreenUpdating = True
    ClearStatus

    If Not gHeadless Then
        MsgBox "FIRMette run complete." & vbCrLf & _
            "Downloaded: " & ok & vbCrLf & _
            "Failed:     " & failed & vbCrLf & vbCrLf & _
            "Folder: " & folder, _
            IIf(failed = 0, vbInformation, vbExclamation), "Download FIRMettes"
    End If
End Sub

Private Function ShouldRunFirmetteRow(ByVal ws As Worksheet, ByVal r As Long, _
        ByVal onlyFailed As Boolean) As Boolean
    If RowIsEmpty(ws, r) Then Exit Function
    If Not HasValidCoords(ws, r) Then Exit Function
    If onlyFailed Then
        ShouldRunFirmetteRow = (InStr(1, CStr(ws.Cells(r, COL_FIRMSTATUS).Value), _
            STATUS_FAILED_PREFIX, vbTextCompare) = 1)
    Else
        ShouldRunFirmetteRow = True
    End If
End Function

' "WO123 DI456 - DR-TEST - SiteName FIRMette.pdf" — each piece omitted
' when blank, so no dangling "WO " or trailing " - " ever appears.
Private Function FirmetteFileName(ByVal wo As String, ByVal di As String, _
        ByVal disaster As String, ByVal siteName As String) As String
    Dim s As String, jobs As String
    jobs = JobIds(wo, di, " ", "WO", "DI")
    If Len(jobs) > 0 Then s = jobs & " - "
    If Len(disaster) > 0 Then s = s & disaster & " - "
    s = s & siteName & " FIRMette.pdf"
    FirmetteFileName = CleanFileName(s)
End Function

' Submit → poll → fetch OutputFile → download. Returns True on success.
Private Function DownloadOneFirmette(ByVal lat As String, ByVal lon As String, _
        ByVal fullPath As String, ByRef errMsg As String) As Boolean
    Dim submitUrl As String, submitJson As String, jobId As String
    Dim status As String, attempt As Long, outputJson As String, pdfUrl As String
    Dim httpErr As String

    submitUrl = REST_FIRMETTE & "/submitJob?input_lat=" & lat & _
        "&input_lon=" & lon & "&Print_Type=FIRMETTE&graphic=PDF&f=pjson"
    submitJson = HttpGetText(submitUrl, httpErr)
    If Len(httpErr) > 0 Then errMsg = "submitJob: " & httpErr: Exit Function
    jobId = FirstString(submitJson, "jobId")
    If Len(jobId) = 0 Then errMsg = "submitJob response missing jobId": Exit Function

    For attempt = 1 To FIRMETTE_POLL_MAX_ATTEMPTS
        status = GetFirmetteJobStatus(jobId, httpErr)
        If Len(httpErr) > 0 Then errMsg = "jobStatus: " & httpErr: Exit Function
        Select Case status
            Case "esriJobSucceeded": Exit For
            Case "esriJobSubmitted", "esriJobExecuting", "esriJobWaiting", "esriJobNew"
                WaitSeconds FIRMETTE_POLL_INTERVAL_SECONDS
            Case Else
                errMsg = "Job did not succeed (status=" & status & ")"
                Exit Function
        End Select
    Next attempt
    If status <> "esriJobSucceeded" Then
        errMsg = "Timed out waiting for FEMA GP job (jobId=" & jobId & ")"
        Exit Function
    End If

    outputJson = HttpGetText(REST_FIRMETTE & "/jobs/" & jobId & "/results/OutputFile?f=pjson", httpErr)
    If Len(httpErr) > 0 Then errMsg = "OutputFile: " & httpErr: Exit Function
    pdfUrl = FirstString(outputJson, "url")
    If Len(pdfUrl) = 0 Then errMsg = "OutputFile response missing url": Exit Function

    DownloadOneFirmette = HttpDownloadPdf(pdfUrl, fullPath, errMsg)
End Function

Private Function GetFirmetteJobStatus(ByVal jobId As String, ByRef errMsg As String) As String
    Dim json As String, httpErr As String
    json = HttpGetText(REST_FIRMETTE & "/jobs/" & jobId & "?f=pjson", httpErr)
    If Len(httpErr) > 0 Then errMsg = httpErr: Exit Function
    GetFirmetteJobStatus = FirstString(json, "jobStatus")
    If Len(GetFirmetteJobStatus) = 0 Then errMsg = "jobStatus response missing jobStatus"
End Function

' ---- Map Pages (Workflow 3, ported from prototype) ------------------------

Public Sub PrepareMapPages()
    Dim wsSites As Worksheet, wsMap As Worksheet, last As Long, r As Long
    Dim pageIdx As Long, pages As Long

    Set wsSites = SitesSheet()
    last = SitesLastRow()
    If last < SITES_FIRST_DATA_ROW Then
        If Not gHeadless Then MsgBox "No site rows found.", vbInformation, "Prepare Map Pages"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Fail

    ' Delete and recreate the MapPages sheet.
    If SheetExists(SH_MAPPAGES) Then ThisWorkbook.Worksheets(SH_MAPPAGES).Delete
    Set wsMap = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsMap.Name = SH_MAPPAGES
    ConfigureMapPageSetup wsMap

    pageIdx = 0
    For r = SITES_FIRST_DATA_ROW To last
        If HasValidCoords(wsSites, r) Then
            CreateMapPage wsMap, wsSites, r, pageIdx
            wsSites.Cells(r, COL_MAPSTATUS).Value = "Map page created"
            pageIdx = pageIdx + 1
        End If
    Next r
    pages = pageIdx

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    wsMap.Activate
    On Error GoTo 0

    If Not gHeadless Then
        MsgBox "Created " & pages & " map page(s) on the '" & SH_MAPPAGES & "' sheet." & vbCrLf & vbCrLf & _
            "Next steps:" & vbCrLf & _
            "1. Paste a screenshot onto each page using Place in Cell." & vbCrLf & _
            "2. Click 'Export Combined Map PDF' when done.", _
            vbInformation, "Map Pages Ready"
    End If
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If gHeadless Then
        Err.Raise Err.Number, "PrepareMapPages", Err.Description
    Else
        MsgBox "Map page creation failed: " & Err.Description, vbCritical, "Prepare Map Pages"
    End If
End Sub

' Append one blank page (no Sites row). Useful when an inspector wants
' an extra overview page after the per-site pages.
Public Sub AddMapPage()
    Dim wsMap As Worksheet, pageIdx As Long

    Application.ScreenUpdating = False
    On Error GoTo Fail
    If Not SheetExists(SH_MAPPAGES) Then
        Set wsMap = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsMap.Name = SH_MAPPAGES
        ConfigureMapPageSetup wsMap
        pageIdx = 0
    Else
        Set wsMap = ThisWorkbook.Worksheets(SH_MAPPAGES)
        ' Pages already present = HPageBreaks.Count + 1 (the last page has no break after it).
        pageIdx = wsMap.HPageBreaks.Count + 1
    End If

    CreateMapPage wsMap, Nothing, 0, pageIdx
    Application.ScreenUpdating = True
    wsMap.Activate
    wsMap.Cells(pageIdx * MAP_ROWS_PER_PAGE + 1, 1).Select

    If Not gHeadless Then
        MsgBox "Added blank page " & (pageIdx + 1) & ".", vbInformation, "Add Page"
    End If
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    If gHeadless Then Err.Raise Err.Number, "AddMapPage", Err.Description Else _
        MsgBox "Add Page failed: " & Err.Description, vbCritical, "Add Page"
End Sub

Private Sub ConfigureMapPageSetup(ByVal wsMap As Worksheet)
    Dim cc As Long
    With wsMap.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperLetter
        .LeftMargin = 0:  .RightMargin = 0
        .TopMargin = 0:   .BottomMargin = 0
        .HeaderMargin = 0: .FooterMargin = 0
        .CenterHorizontally = True
        .CenterVertically = True
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintGridlines = False
        .PrintHeadings = False
    End With
    For cc = 1 To MAP_COLS_WIDE
        wsMap.Columns(cc).ColumnWidth = 10.5
    Next cc
End Sub

' Lay out one page (4 merged rows x 13 cols) plus the WO/DI/applicant textbox.
' If wsSites is Nothing the page is blank (used by AddMapPage).
Private Sub CreateMapPage(ByVal wsMap As Worksheet, ByVal wsSites As Worksheet, _
        ByVal siteRow As Long, ByVal pageIdx As Long)
    Dim startRow As Long, rr As Long, rowH As Double, pageTopPts As Double
    Dim placeholderTxt As String, sNum As String, sNam As String
    Dim txtBox As Shape, txtContent As String, firstLineLen As Long

    rowH = MAP_PAGE_HEIGHT_PTS / MAP_ROWS_PER_PAGE
    startRow = pageIdx * MAP_ROWS_PER_PAGE + 1

    For rr = startRow To startRow + MAP_ROWS_PER_PAGE - 1
        wsMap.Rows(rr).RowHeight = rowH
    Next rr

    wsMap.Range(wsMap.Cells(startRow, 1), wsMap.Cells(startRow + MAP_ROWS_PER_PAGE - 1, MAP_COLS_WIDE)).Merge

    ' Placeholder text in the merged cell.
    If wsSites Is Nothing Then
        placeholderTxt = "New Page" & vbCrLf & vbCrLf & "Paste screenshot here (Place in Cell)"
    Else
        sNum = Trim$(CStr(wsSites.Cells(siteRow, COL_SITENO).Value))
        sNam = Trim$(CStr(wsSites.Cells(siteRow, COL_SITENAME).Value))
        If Len(sNum) > 0 Then
            placeholderTxt = "Site " & sNum & " " & sNam
        Else
            placeholderTxt = sNam
        End If
        placeholderTxt = placeholderTxt & vbCrLf & vbCrLf & "Paste screenshot here (Place in Cell)"
    End If
    With wsMap.Cells(startRow, 1)
        .Value = placeholderTxt
        .Font.Color = RGB(200, 200, 200)
        .Font.Size = 18
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' WO/DI/Applicant textbox in the top-left of the page area.
    pageTopPts = 0
    For rr = 1 To startRow - 1
        pageTopPts = pageTopPts + wsMap.Rows(rr).Height
    Next rr

    If wsSites Is Nothing Then
        txtContent = "[Edit this textbox]"
    Else
        txtContent = BuildMapTextboxString(wsSites, siteRow)
    End If

    Set txtBox = wsMap.Shapes.AddTextbox(msoTextOrientationHorizontal, _
        5, pageTopPts + 5, MAP_TEXTBOX_WIDTH, MAP_TEXTBOX_HEIGHT)
    With txtBox
        .Name = "Textbox_Page_" & CStr(pageIdx + 1)
        With .TextFrame2
            .WordWrap = msoTrue
            .MarginLeft = 5:  .MarginRight = 5
            .MarginTop = 3:    .MarginBottom = 3
            .AutoSize = msoAutoSizeNone
        End With
        With .TextFrame
            .Characters.Text = txtContent
            .Characters.Font.Name = "Segoe UI"
            .Characters.Font.Size = 8
            .Characters.Font.Color = RGB(0, 0, 0)
        End With
        firstLineLen = InStr(1, txtContent, vbLf) - 1
        If firstLineLen > 0 Then .TextFrame.Characters(1, firstLineLen).Font.Bold = True
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
            .Transparency = 0.2
        End With
        With .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(100, 100, 100)
            .Weight = 0.5
        End With
    End With

    ' Page break AFTER this page so the next CreateMapPage call lands on a new sheet page.
    On Error Resume Next
    wsMap.HPageBreaks.Add Before:=wsMap.Rows(startRow + MAP_ROWS_PER_PAGE)
    On Error GoTo 0
End Sub

' Build the WO #... / Applicant / Site N / lat,lon / Cat ... / desc textbox stamp.
' Reads WO/DI from the row first (per-row override), then Setup.
Private Function BuildMapTextboxString(ByVal wsSites As Worksheet, ByVal r As Long) As String
    Dim wo As String, di As String, applicant As String
    Dim sNum As String, sNam As String, lat As String, lon As String
    Dim cat As String, desc As String, siteLine As String, catLine As String

    wo = Trim$(CStr(wsSites.Cells(r, COL_WO).Value))
    di = Trim$(CStr(wsSites.Cells(r, COL_DI).Value))
    If Len(wo) = 0 Then wo = SetupValue(NR_WO)
    If Len(di) = 0 Then di = SetupValue(NR_DI)
    applicant = SetupValue(NR_APPLICANT)

    sNum = Trim$(CStr(wsSites.Cells(r, COL_SITENO).Value))
    sNam = Trim$(CStr(wsSites.Cells(r, COL_SITENAME).Value))
    lat = Format$(wsSites.Cells(r, COL_LAT).Value, "0.00000")
    lon = Format$(wsSites.Cells(r, COL_LON).Value, "0.00000")
    cat = Trim$(CStr(wsSites.Cells(r, COL_CATEGORY).Value))
    desc = Trim$(CStr(wsSites.Cells(r, COL_DESC).Value))

    If Len(sNum) > 0 Then
        siteLine = "Site " & sNum & " " & sNam
    Else
        siteLine = sNam
    End If

    If Len(cat) > 0 And Len(desc) > 0 Then
        catLine = "Cat " & cat & ", " & desc
    ElseIf Len(cat) > 0 Then
        catLine = "Cat " & cat
    Else
        catLine = desc
    End If

    Dim woDiLine As String
    woDiLine = JobIds(wo, di, " ", "WO #", "DI #")     ' "" when both blank

    ' Build the stamp top-down, skipping the WO/DI line entirely when the
    ' inspector didn't enter either ID. The applicant line still anchors
    ' the rest of the textbox.
    Dim costs As String, workComp As String
    costs = Trim$(CStr(wsSites.Cells(r, COL_COSTS).Value))
    workComp = Trim$(CStr(wsSites.Cells(r, COL_WORKCOMP).Value))

    If Len(woDiLine) > 0 Then BuildMapTextboxString = woDiLine & vbLf
    BuildMapTextboxString = BuildMapTextboxString & _
        applicant & vbLf & _
        siteLine & vbLf & _
        lat & ", " & lon
    If Len(catLine) > 0 Then BuildMapTextboxString = BuildMapTextboxString & vbLf & catLine
    ' Two optional money/progress lines, only emitted when populated.
    If Len(costs) > 0 Then BuildMapTextboxString = BuildMapTextboxString & vbLf & "Cost: " & costs
    If Len(workComp) > 0 Then BuildMapTextboxString = BuildMapTextboxString & vbLf & "Work: " & workComp
End Function

Public Sub ExportCombinedMapPdf()
    Dim wsMap As Worksheet, folder As String, fileName As String, fullPath As String
    Dim disaster As String, wo As String, di As String

    If Not SheetExists(SH_MAPPAGES) Then
        If Not gHeadless Then MsgBox "No '" & SH_MAPPAGES & "' sheet. Click 'Prepare Map Pages' first.", vbExclamation, "Export Map PDF"
        Exit Sub
    End If
    Set wsMap = ThisWorkbook.Worksheets(SH_MAPPAGES)

    folder = ResolveOutputFolder()
    If Not EnsureFolderExists(folder) Then
        If Not gHeadless Then MsgBox "Could not create the output folder:" & vbCrLf & folder, vbExclamation, "Export Map PDF"
        Exit Sub
    End If
    disaster = SetupValue(NR_DISASTER)
    wo = SetupValue(NR_WO)
    di = SetupValue(NR_DI)

    Dim jobs As String
    jobs = JobIds(wo, di, " ", "WO", "DI")
    fileName = ""
    If Len(jobs) > 0 Then fileName = jobs & " - "
    If Len(disaster) > 0 Then fileName = fileName & disaster & " - "
    fileName = fileName & "Location Map.pdf"
    fileName = CleanFileName(fileName)
    fullPath = folder & fileName

    ' Bring text boxes to the front in case the inspector's pasted screenshot covered them.
    EnsureTextboxesOnTop wsMap

    On Error GoTo Fail
    wsMap.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fullPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
    On Error GoTo 0

    If Not gHeadless Then MsgBox "Combined map PDF exported:" & vbCrLf & fullPath, vbInformation, "Export Map PDF"
    Exit Sub
Fail:
    If gHeadless Then Err.Raise Err.Number, "ExportCombinedMapPdf", Err.Description Else _
        MsgBox "Export failed: " & Err.Description, vbCritical, "Export Map PDF"
End Sub

Private Sub EnsureTextboxesOnTop(ByVal wsMap As Worksheet)
    Dim shp As Shape
    For Each shp In wsMap.Shapes
        If shp.Type = msoTextBox Then shp.ZOrder msoBringToFront
    Next shp
End Sub

' ---- KML export (F10) -----------------------------------------------------

Public Sub ExportSitesToKML()
    Dim file As String, n As Long, dialogTitle As String
    dialogTitle = "Export KML"
    If Not WriteSitesKml(file, n, dialogTitle) Then Exit Sub
    If Not gHeadless Then
        Dim q As String: q = Chr$(34)
        Shell "cmd /c start " & q & q & " " & q & file & q, vbNormalFocus
        MsgBox "Exported " & n & " point(s) to:" & vbCrLf & file, vbInformation, dialogTitle
    End If
End Sub

' Build the Sites KML and open the inspector's AGOL webmap so they can
' drag the file onto it (AGOL Map Viewer's "Add Layer from File" pattern).
' Requires NR_AGOLMAP to be set on Setup.
Public Sub SendSitesToAgolMap()
    Dim agol As String: agol = SetupValue(NR_AGOLMAP)
    If Len(agol) = 0 Then
        If Not gHeadless Then MsgBox _
            "No AGOL Webmap URL set. Paste your map's URL on the Setup sheet " & _
            "first, then click this again.", vbExclamation, "Send to AGOL Map"
        Exit Sub
    End If

    Dim file As String, n As Long, dialogTitle As String
    dialogTitle = "Send to AGOL Map"
    If Not WriteSitesKml(file, n, dialogTitle) Then Exit Sub

    If Not gHeadless Then
        ' Open the inspector's AGOL webmap in their default browser.
        On Error Resume Next
        ThisWorkbook.FollowHyperlink Address:=agol, NewWindow:=False
        On Error GoTo 0
        ' And open Explorer at the KML's folder so it's a single drag-drop
        ' onto the AGOL Map Viewer window.
        Dim q As String: q = Chr$(34)
        Shell "explorer.exe /select," & q & file & q, vbNormalFocus
        MsgBox "Exported " & n & " point(s) to:" & vbCrLf & file & vbCrLf & vbCrLf & _
            "Your AGOL webmap should now be open in the browser." & vbCrLf & _
            "Drag the highlighted KML file from Explorer onto the Map Viewer window" & vbCrLf & _
            "to add the sites as a new layer.", _
            vbInformation, dialogTitle
    End If
End Sub

' Shared KML builder: writes the file to the resolved output folder, sets
' filePath + featureCount on success, returns False on any failure (with a
' MsgBox already shown if gHeadless is False).
Private Function WriteSitesKml(ByRef filePath As String, ByRef featureCount As Long, _
        ByVal dialogTitle As String) As Boolean
    Dim ws As Worksheet, last As Long, r As Long, kml As String, n As Long
    Set ws = SitesSheet()
    last = SitesLastRow()
    If last < SITES_FIRST_DATA_ROW Then
        If Not gHeadless Then MsgBox "No site rows to export.", vbInformation, dialogTitle
        Exit Function
    End If

    kml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
        "<kml xmlns=""http://www.opengis.net/kml/2.2""><Document>" & vbCrLf & _
        "<name>RoadReviewer Sites</name>" & vbCrLf & _
        "<open>1</open>" & vbCrLf & _
        KmlPinStyles()
    ' Bucket rows by Category into <Folder> elements so the inspector can
    ' expand / collapse by category in Google Earth's left-side tree (and
    ' AGOL Map Viewer's contents panel when the KML is loaded as a layer).
    kml = kml & BuildCategoryFolders(ws, last, n)
    kml = kml & "</Document></kml>"

    If n = 0 Then
        If Not gHeadless Then MsgBox "No rows have valid coordinates to export.", vbInformation, dialogTitle
        Exit Function
    End If

    Dim folder As String, file As String
    folder = ResolveOutputFolder()
    If Not EnsureFolderExists(folder) Then
        If Not gHeadless Then MsgBox "Could not create the output folder:" & vbCrLf & folder, vbExclamation, dialogTitle
        Exit Function
    End If
    file = folder & "RoadReviewer Sites.kml"

    If Not WriteTextFile(file, kml) Then
        If Not gHeadless Then MsgBox "Could not write the KML file.", vbExclamation, dialogTitle
        Exit Function
    End If

    filePath = file
    featureCount = n
    WriteSitesKml = True
End Function

' Walk the Sites table twice:
'   pass 1 — collect every unique Category value (trimmed, case-preserved
'            but case-insensitive grouping) into the order it first
'            appears, so the folder list is predictable
'   pass 2 — for each Category bucket, emit a <Folder> with all its rows
' Rows with a blank Category fall into a "(no category)" folder so they
' aren't lost. Updates featureCount with the number of placemarks emitted.
Private Function BuildCategoryFolders(ByVal ws As Worksheet, ByVal last As Long, _
        ByRef featureCount As Long) As String
    Dim r As Long, cat As String, key As String
    Dim order As Object, rows As Object   ' Scripting.Dictionary
    Set order = CreateObject("Scripting.Dictionary")    ' key -> display label
    Set rows = CreateObject("Scripting.Dictionary")     ' key -> "|row1|row2|..."

    For r = SITES_FIRST_DATA_ROW To last
        If HasValidCoords(ws, r) Then
            cat = Trim$(CStr(ws.Cells(r, COL_CATEGORY).Value))
            If Len(cat) = 0 Then
                key = "__nocat__"
                If Not order.Exists(key) Then order.Add key, "(no category)"
            Else
                key = LCase$(cat)
                If Not order.Exists(key) Then order.Add key, "Category " & cat
            End If
            If Not rows.Exists(key) Then rows.Add key, ""
            rows(key) = rows(key) & "|" & r
        End If
    Next r

    Dim out As String, k As Variant, parts() As String, p As Variant
    For Each k In order.Keys
        out = out & "<Folder><name>" & XmlEscape(CStr(order(k))) & "</name>" & _
            "<open>1</open>" & vbCrLf
        parts = Split(Mid$(CStr(rows(k)), 2), "|")  ' drop leading | then split
        For Each p In parts
            out = out & PlacemarkXml(ws, CLng(p))
            featureCount = featureCount + 1
        Next p
        out = out & "</Folder>" & vbCrLf
    Next k
    BuildCategoryFolders = out
End Function

' KML style block. Three named styles for the three Federal Aid Status
' buckets the classifier produces. Rows that haven't been classified
' (blank Federal Aid Status cell) fall back to the "review" style so the
' inspector can still see them on the map as something-to-look-at.
' Colors are KML's ABGR hex (alpha-blue-green-red, NOT RGBA).
Private Function KmlPinStyles() As String
    KmlPinStyles = _
        "<Style id=""fedAid""><IconStyle><color>ff0000ff</color>" & _
        "<Icon><href>http://maps.google.com/mapfiles/kml/paddle/red-circle.png</href></Icon></IconStyle></Style>" & vbCrLf & _
        "<Style id=""nonFedAid""><IconStyle><color>ff00ff00</color>" & _
        "<Icon><href>http://maps.google.com/mapfiles/kml/paddle/grn-circle.png</href></Icon></IconStyle></Style>" & vbCrLf & _
        "<Style id=""review""><IconStyle><color>ffffffff</color>" & _
        "<Icon><href>http://maps.google.com/mapfiles/kml/paddle/ltblu-circle.png</href></Icon></IconStyle></Style>" & vbCrLf
End Function

' Map the Federal Aid Status cell text to one of the style ids defined
' in KmlPinStyles. Blank / Failed / out-of-state rows fall back to
' "review" (light blue) so unclassified points still show on the map
' as something the inspector should look at.
Private Function PinStyleId(ByVal status As String) As String
    Dim s As String: s = LCase$(Trim$(status))
    If Len(s) = 0 Then PinStyleId = "review": Exit Function
    If Left$(s, 15) = "non-federal aid" Then PinStyleId = "nonFedAid": Exit Function
    If Left$(s, 11) = "federal aid" Then PinStyleId = "fedAid": Exit Function
    If Left$(s, 6) = "review" Then PinStyleId = "review": Exit Function
    If InStr(s, "no road segment") > 0 Then PinStyleId = "review": Exit Function
End Function

Private Function PlacemarkXml(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim nm As String, desc As String, lat As String, lon As String
    Dim status As String, styleId As String, styleTag As String
    nm = XmlEscape(CStr(ws.Cells(r, COL_SITENAME).Value))
    If Len(nm) = 0 Then nm = "Site row " & r
    status = CStr(ws.Cells(r, COL_ELIGIBILITY).Value)
    desc = XmlEscape(BuildDescBlock(ws, r, status))
    lat = InvariantNum(ws.Cells(r, COL_LAT).Value)
    lon = InvariantNum(ws.Cells(r, COL_LON).Value)
    styleId = PinStyleId(status)
    If Len(styleId) > 0 Then styleTag = "<styleUrl>#" & styleId & "</styleUrl>"
    PlacemarkXml = "<Placemark><name>" & nm & "</name>" & styleTag & _
        IIf(Len(desc) > 0, "<description>" & desc & "</description>", "") & _
        "<Point><coordinates>" & lon & "," & lat & ",0</coordinates></Point></Placemark>" & vbCrLf
End Function

' Pipe-separated description block used in KML <description>. Each field
' is omitted when blank, so we never produce dangling " | " separators.
' Order: Description, Costs, Work Completion, Federal Aid Status. Status
' goes last so it reads like a tag suffix.
Private Function BuildDescBlock(ByVal ws As Worksheet, ByVal r As Long, _
        ByVal status As String) As String
    Dim parts() As String, n As Integer, v As String
    ReDim parts(3)
    n = 0
    v = Trim$(CStr(ws.Cells(r, COL_DESC).Value))
    If Len(v) > 0 Then parts(n) = v: n = n + 1
    v = Trim$(CStr(ws.Cells(r, COL_COSTS).Value))
    If Len(v) > 0 Then parts(n) = "Cost: " & v: n = n + 1
    v = Trim$(CStr(ws.Cells(r, COL_WORKCOMP).Value))
    If Len(v) > 0 Then parts(n) = "Work: " & v: n = n + 1
    v = Trim$(status)
    If Len(v) > 0 Then parts(n) = v: n = n + 1
    If n = 0 Then Exit Function
    ReDim Preserve parts(n - 1)
    BuildDescBlock = Join(parts, " | ")
End Function

Private Function XmlEscape(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    XmlEscape = s
End Function

Private Function WriteTextFile(ByVal path As String, ByVal content As String) As Boolean
    Dim fso As Object, ts As Object
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(path, True, False)   ' overwrite, ANSI (ASCII-clean KML)
    ts.Write content
    ts.Close
    WriteTextFile = True
    Exit Function
Fail:
    WriteTextFile = False
End Function
