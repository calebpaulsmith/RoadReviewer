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

' On-sheet MapPages control shapes (buttons + notes parked right of the print
' grid). Shared prefix so they're easy to hide as a group during PDF export.
Private Const MAP_CTRL_PREFIX As String = "MapCtrl_"

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

' Effective output folder: the explicit Start Here / Map Pages value if set,
' else a per-product subfolder NEXT TO THE WORKBOOK (PR #37, user direction):
' "<workbook dir>\RR Output\" (standard) or "<workbook dir>\SI Tool Output\"
' (inspector), created on first use. The §8.9 OneDrive probe is only the last
' resort for a workbook whose folder genuinely can't be determined (never
' saved, or an https path with no local OneDrive mapping).
Public Function ResolveOutputFolder() As String
    Dim v As String
    v = SetupValue(NR_OUTFOLDER)
    If Len(v) = 0 Then
        Dim base As String
        base = WorkbookFolder()
        If Len(base) > 0 Then
            v = base & "\" & OutputSubfolderName()
        Else
            v = DefaultOutputFolder()
        End If
    End If
    If Right$(v, 1) <> "\" Then v = v & "\"
    ResolveOutputFolder = v
End Function

' The per-product output subfolder created next to the workbook.
Public Function OutputSubfolderName() As String
    If ProductIsInspector() Then
        OutputSubfolderName = "SI Tool Output"
    Else
        OutputSubfolderName = "RR Output"
    End If
End Function

' The folder this .xlsm lives in, or "" when there isn't a usable one.
' A OneDrive/SharePoint-synced workbook reports an https:// URL here (that is
' how "exports save next to the file" silently degraded to the §8.9 probe -
' the bug the user hit); map it back to the locally synced path when possible.
Private Function WorkbookFolder() As String
    Dim p As String
    p = ThisWorkbook.Path
    If Len(p) = 0 Then Exit Function                     ' never saved
    If LCase$(Left$(p, 4)) = "http" Then p = OneDriveLocalFolder(p)
    WorkbookFolder = p
End Function

' Map an https://... OneDrive/SharePoint folder URL to its locally synced
' path by testing progressively shorter URL tails under each OneDrive
' env-var root until this workbook's file is found on disk. Returns "" when
' no mapping resolves (the caller then falls back to the §8.9 probe).
Private Function OneDriveLocalFolder(ByVal urlFolder As String) As String
    Dim bases As Variant, parts() As String
    Dim b As Long, i As Long, j As Long, tail As String, cand As String
    ' The longest candidate tails still contain "https:" etc. - Dir$ raises
    ' error 52 on those instead of returning ""; treat any error as no-match.
    On Error Resume Next
    bases = Array(Environ$("OneDriveCommercial"), Environ$("OneDriveConsumer"), Environ$("OneDrive"))
    parts = Split(Replace(urlFolder, "%20", " "), "/")
    For b = LBound(bases) To UBound(bases)
        If Len(bases(b)) > 0 Then
            For i = LBound(parts) To UBound(parts)
                tail = ""
                For j = i To UBound(parts)
                    If Len(parts(j)) > 0 Then tail = tail & "\" & parts(j)
                Next j
                If Len(tail) > 0 Then
                    cand = bases(b) & tail
                    If Len(Dir$(cand & "\" & ThisWorkbook.Name)) > 0 Then
                        OneDriveLocalFolder = cand
                        Exit Function
                    End If
                End If
            Next i
        End If
    Next b
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
    DefaultOutputFolder = base & "\Desktop\Scripts\RoadReviewer\" & _
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

    ' Leave ScreenUpdating on and stay on Sites so the inspector watches
    ' FIRMette Status fill in row by row - each row is a submitJob/poll/
    ' download round trip against FEMA's GP service, so redraw cost here
    ' is negligible next to the network wait.
    ws.Activate
    For r = SITES_FIRST_DATA_ROW To last
        If Not ShouldRunFirmetteRow(ws, r, onlyFailed) Then GoTo NextRow
        processed = processed + 1
        siteName = CStr(ws.Cells(r, COL_SITENAME).Value)
        If Len(Trim$(siteName)) = 0 Then siteName = "row" & r
        SetStatus "Downloading FIRMette " & processed & " of " & total & " - " & siteName
        ws.Cells(r, COL_FIRMSTATUS).Select
        DoEvents

        lat = InvariantNum(ws.Cells(r, COL_LAT).Value)
        lon = InvariantNum(ws.Cells(r, COL_LON).Value)
        fileName = FirmetteFileName(siteName)
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
    ClearStatus
    If ok > 0 Then SurfaceFolder folder

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

' "WO123 DI5 - DR-4882-IN - SiteName FIRMette.pdf" - the shared JobFileStem
' (WO/DI/Disaster-State, or a date-time when all blank) plus the per-site name.
Private Function FirmetteFileName(ByVal siteName As String) As String
    FirmetteFileName = CleanFileName(JobFileStem() & " - " & siteName & " FIRMette.pdf")
End Function

' Live example of what DownloadFirmettes will name its files, built from the
' FIRST site row and the current job info. Used as a worksheet formula
' (=FirmettePreview()) on Start Here's FIRMette section, so the inspector can
' see the effect of the WO/DI/Disaster fields - which now live over on MapPages
' - without leaving Start Here. Application.Volatile so it recalculates when
' any of those cells change.
Public Function FirmettePreview() As String
    Application.Volatile
    Dim ws As Worksheet, last As Long, r As Long, siteName As String

    On Error GoTo Fallback
    Set ws = SitesSheet()
    If ws Is Nothing Then GoTo Fallback
    last = SitesLastRow()

    For r = SITES_FIRST_DATA_ROW To last
        If HasValidCoords(ws, r) Then
            siteName = Trim$(CStr(ws.Cells(r, COL_SITENAME).Value))
            If Len(siteName) = 0 Then siteName = "Site " & CStr(r - SITES_FIRST_DATA_ROW + 1)
            FirmettePreview = FirmetteFileName(siteName)
            Exit Function
        End If
    Next r

Fallback:
    ' No site rows yet - show the shape of the name with a placeholder site.
    ' NB: no angle brackets - CleanFileName turns them into underscores.
    FirmettePreview = FirmetteFileName("(site name)")
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

' Create the MapPages sheet if it doesn't exist and (re)write its header band:
' the job inputs + the tools panel. Called by BuildWorkbook, so the sheet - and
' the JobWO/JobDI/... named ranges the inspector product parks on it - ALWAYS
' exist, even before any page has been built.
'
' Idempotent and non-destructive: existing pages and existing job values are
' left alone, so Repair Layout and a re-run of Prepare Map Pages are both safe.
Public Sub EnsureMapPagesSheet()
    Dim wsMap As Worksheet, isNew As Boolean

    If SheetExists(SH_MAPPAGES) Then
        Set wsMap = ThisWorkbook.Worksheets(SH_MAPPAGES)
    Else
        Set wsMap = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsMap.Name = SH_MAPPAGES
        isNew = True
    End If

    ConfigureMapPageSetup wsMap
    BuildMapHeaderBand wsMap
    AddMapPageControls wsMap
    SetMapPrintArea wsMap
    wsMap.Tab.Color = RGB(60, 60, 60)
End Sub

' Reveal MapPages. On the standard product it ships HIDDEN (map pages are an
' opt-in there); any user-triggered map action unhides it. Idempotent.
Public Sub ShowMapPages()
    On Error Resume Next
    If SheetExists(SH_MAPPAGES) Then ThisWorkbook.Worksheets(SH_MAPPAGES).Visible = xlSheetVisible
    On Error GoTo 0
End Sub

' "Back to Map Pages" on the inspector's hidden Tools sheet - reveal + jump to
' the map workspace where the whole map/FIRMette workflow lives.
Public Sub GoToMapPages()
    EnsureMapPagesSheet
    ShowMapPages
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_MAPPAGES).Activate
    On Error GoTo 0
End Sub

' "Exports & other tools" on the Map Pages band (inspector) - reveal + jump to
' the SH_START utility sheet (hand-off exports, optional FHWA, Repair/Reset),
' which ships hidden so the inspector opens straight onto Map Pages.
Public Sub GoToOtherTools()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(StartSheetName())
    If Not ws Is Nothing Then
        ws.Visible = xlSheetVisible
        ws.Activate
    End If
    On Error GoTo 0
End Sub

' Restrict printing to the page columns, from the first page row down - so the
' header band (job inputs) and the tools panel never land in the PDF. Without a
' PrintArea the panel cells would spill onto extra pages.
Private Sub SetMapPrintArea(ByVal wsMap As Worksheet)
    Dim lastRow As Long, savedAlerts As Boolean
    lastRow = MAP_FIRST_PAGE_ROW + (MapPageCount(wsMap) * MAP_ROWS_PER_PAGE) - 1
    If lastRow < MAP_FIRST_PAGE_ROW Then lastRow = MAP_FIRST_PAGE_ROW
    On Error Resume Next
    ' With ONE site the print area is a single MERGED page cell, and Excel then
    ' pops "You've selected a single cell for the print area" on the assignment
    ' when alerts are on (PrepareMapPages re-enables them before Export runs).
    ' OK is the right answer - the merged block IS the page - so suppress the
    ' prompt; it hung every headless 1-site export and confused real users.
    savedAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wsMap.PageSetup.PrintArea = wsMap.Range( _
        wsMap.Cells(MAP_FIRST_PAGE_ROW, 1), wsMap.Cells(lastRow, MAP_COLS_WIDE)).Address
    Application.DisplayAlerts = savedAlerts
    On Error GoTo 0
End Sub

' Pages present = one "Textbox_Page_N" per page. Counting the shapes is exact;
' the old HPageBreaks.Count+1 trick breaks once a PrintArea is set.
Public Function MapPageCount(ByVal wsMap As Worksheet) As Long
    Dim shp As Shape, n As Long
    For Each shp In wsMap.Shapes
        If Left$(shp.Name, Len("Textbox_Page_")) = "Textbox_Page_" Then n = n + 1
    Next shp
    MapPageCount = n
End Function

' Advanced-options step 1 (also runnable on its own). The one-click
' CreateMapPagesPdf chains PreparePagesCore -> FetchImageryCore ->
' ExportMapPdfCore with a single summary message at the end.
Public Sub PrepareMapPages()
    Dim pages As Long
    pages = PreparePagesCore()
    If pages > 0 And Not gHeadless Then
        MsgBox "Created " & pages & " map page(s) on the '" & SH_MAPPAGES & "' sheet." & vbCrLf & vbCrLf & _
            "Next steps (Advanced options):" & vbCrLf & _
            "2. Click 'Fetch Imagery' - an aerial image is downloaded and placed on every page" & vbCrLf & _
            "   automatically (or use the manual Google Earth screenshot buttons instead)." & vbCrLf & _
            "3. Click 'Export PDF' when done.", _
            vbInformation, "Map Pages Ready"
    End If
End Sub

' Builds one page per Sites row with valid coordinates. Returns the page count
' (0 when there was nothing to do or the build failed non-headless).
Public Function PreparePagesCore() As Long
    Dim wsSites As Worksheet, wsMap As Worksheet, last As Long, r As Long
    Dim pageIdx As Long

    Set wsSites = SitesSheet()
    last = SitesLastRow()
    If last < SITES_FIRST_DATA_ROW Then
        If Not gHeadless Then MsgBox "No site rows found.", vbInformation, "Prepare Map Pages"
        Exit Function
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Fail

    ' The sheet is NO LONGER deleted and recreated - it now holds the job inputs
    ' and their named ranges. Only the pages themselves are cleared and rebuilt.
    EnsureMapPagesSheet
    ShowMapPages                 ' opt-in reveal on the standard product
    Set wsMap = ThisWorkbook.Worksheets(SH_MAPPAGES)
    ClearMapPages wsMap

    pageIdx = 0
    For r = SITES_FIRST_DATA_ROW To last
        If HasValidCoords(wsSites, r) Then
            CreateMapPage wsMap, wsSites, r, pageIdx
            wsSites.Cells(r, COL_MAPSTATUS).Value = "Map page created"
            pageIdx = pageIdx + 1
        End If
    Next r
    PreparePagesCore = pageIdx

    AddMapPageControls wsMap
    SetMapPrintArea wsMap

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    wsMap.Activate
    On Error GoTo 0
    Exit Function
Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    PreparePagesCore = 0
    If gHeadless Then
        Err.Raise Err.Number, "PrepareMapPages", Err.Description
    Else
        MsgBox "Map page creation failed: " & Err.Description, vbCritical, "Prepare Map Pages"
    End If
End Function

' Wipe the pages (shapes + merged grid) WITHOUT touching the header band, the
' job inputs or the tools panel.
Private Sub ClearMapPages(ByVal wsMap As Worksheet)
    Dim shp As Shape, i As Long, nm As String
    For i = wsMap.Shapes.Count To 1 Step -1
        Set shp = wsMap.Shapes(i)
        nm = shp.Name
        If Left$(nm, Len("Textbox_Page_")) = "Textbox_Page_" _
            Or Left$(nm, Len(MAP_PICKBTN_PREFIX)) = MAP_PICKBTN_PREFIX _
            Or Left$(nm, Len(MAP_PIN_PREFIX)) = MAP_PIN_PREFIX _
            Or Left$(nm, Len(MAP_ATTR_PREFIX)) = MAP_ATTR_PREFIX _
            Or shp.Type = msoPicture Then
            shp.Delete
        End If
    Next i

    Dim lastRow As Long
    lastRow = wsMap.Cells(wsMap.Rows.Count, 1).End(xlUp).Row
    If lastRow < MAP_FIRST_PAGE_ROW Then lastRow = MAP_FIRST_PAGE_ROW
    With wsMap.Range(wsMap.Cells(MAP_FIRST_PAGE_ROW, 1), wsMap.Cells(lastRow + MAP_ROWS_PER_PAGE, MAP_COLS_WIDE))
        .UnMerge
        .Clear
    End With
End Sub

' Rewrite every page's stamp from the CURRENT Sites row + job values, leaving
' the pasted images alone. The stamp is static text baked into a shape when the
' page is built, so without this, editing WO/DI (or the Applicant) after the
' fact would change nothing on the pages. Each stamp remembers its Sites row in
' the shape's AlternativeText.
Public Sub UpdateMapStamps()
    Dim wsMap As Worksheet, wsSites As Worksheet, shp As Shape
    Dim siteRow As Long, txt As String, n As Long, firstLineLen As Long

    If Not SheetExists(SH_MAPPAGES) Then
        If Not gHeadless Then MsgBox "No '" & SH_MAPPAGES & "' sheet yet. Click 'Create Map Pages PDF' first.", _
            vbExclamation, "Update Stamps"
        Exit Sub
    End If
    Set wsMap = ThisWorkbook.Worksheets(SH_MAPPAGES)
    Set wsSites = SitesSheet()
    ShowMapPages

    Application.ScreenUpdating = False
    On Error GoTo Fail
    For Each shp In wsMap.Shapes
        If Left$(shp.Name, Len("Textbox_Page_")) = "Textbox_Page_" Then
            siteRow = 0
            On Error Resume Next
            siteRow = CLng(Val(shp.AlternativeText))
            On Error GoTo Fail
            If siteRow >= SITES_FIRST_DATA_ROW Then
                txt = BuildMapTextboxString(wsSites, siteRow)
                With shp.TextFrame
                    .Characters.Text = txt
                    .Characters.Font.Name = "Segoe UI"
                    .Characters.Font.Size = MAP_STAMP_FONT
                    .Characters.Font.Color = RGB(0, 0, 0)
                End With
                firstLineLen = InStr(1, txt, vbLf) - 1
                If firstLineLen > 0 Then shp.TextFrame.Characters(1, firstLineLen).Font.Bold = True
                n = n + 1
            End If
        End If
    Next shp
    Application.ScreenUpdating = True
    On Error GoTo 0

    If Not gHeadless Then MsgBox "Updated the stamp on " & n & " page(s).", vbInformation, "Update Stamps"
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    If gHeadless Then Err.Raise Err.Number, "UpdateMapStamps", Err.Description Else _
        MsgBox "Update Stamps failed: " & Err.Description, vbCritical, "Update Stamps"
End Sub

' Append one blank page (no Sites row). Useful when an inspector wants
' an extra overview page after the per-site pages.
Public Sub AddMapPage()
    Dim wsMap As Worksheet, pageIdx As Long

    Application.ScreenUpdating = False
    On Error GoTo Fail
    EnsureMapPagesSheet
    ShowMapPages                 ' opt-in reveal on the standard product
    Set wsMap = ThisWorkbook.Worksheets(SH_MAPPAGES)
    ' Count the page stamps, not HPageBreaks - a PrintArea is set now, which
    ' makes the old HPageBreaks.Count+1 trick wrong.
    pageIdx = MapPageCount(wsMap)

    CreateMapPage wsMap, Nothing, 0, pageIdx
    AddMapPageControls wsMap
    SetMapPrintArea wsMap
    Application.ScreenUpdating = True
    wsMap.Activate
    wsMap.Cells(MAP_FIRST_PAGE_ROW + pageIdx * MAP_ROWS_PER_PAGE, 1).Select

    If Not gHeadless Then
        MsgBox "Added blank page " & (pageIdx + 1) & ".", vbInformation, "Add Page"
    End If
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    If gHeadless Then Err.Raise Err.Number, "AddMapPage", Err.Description Else _
        MsgBox "Add Page failed: " & Err.Description, vbCritical, "Add Page"
End Sub

' The header band above the pages. Two cell-based regions (job info + the
' precedence note); the buttons themselves are absolute-positioned shapes added
' by AddMapPageControls, so they're free of the tall page-row heights. Job cells
' carry the JobWO/... named ranges on BOTH products now (the standard product's
' MapPages is hidden until the user opts into map pages, then has the full
' workflow). Values already typed are preserved on a rebuild.
Private Sub BuildMapHeaderBand(ByVal wsMap As Worksheet)
    Dim r As Long

    For r = 1 To MAP_HEADER_ROWS
        wsMap.Rows(r).RowHeight = MAP_HEADER_ROW_HEIGHT
    Next r
    wsMap.Range(wsMap.Cells(1, 1), wsMap.Cells(MAP_HEADER_ROWS, MAP_COLS_WIDE)).Interior.Color = RGB(245, 247, 249)

    ' Section labels (the buttons that go with them are added later, as shapes).
    HeaderLabel wsMap, MAP_JOB_FIRST_ROW - 1, MAP_JOB_LABEL_COL, "Job info  (stamped onto every map page + used in file names)"

    Dim rr As Long
    rr = MAP_JOB_FIRST_ROW
    MapJobField wsMap, rr, "Work Order (WO #)", NR_WO:                 rr = rr + 1
    MapJobField wsMap, rr, "Impact (DI #)", NR_DI:                     rr = rr + 1
    MapJobField wsMap, rr, "Disaster (e.g. 4882)", NR_DISASTER:        rr = rr + 1
    ' State lives HERE for the inspector (Map Pages is its landing) - it drives
    ' road classification AND is appended to the disaster in file names
    ' ("DR-4882-IN"). On the standard product State stays on Start Here (its hub),
    ' so it's not repeated here.
    If ProductIsInspector() Then
        MapJobField wsMap, rr, "State", NR_STATE
        AddMapStateValidation wsMap.Cells(rr, MAP_JOB_VALUE_COL)
        rr = rr + 1
    End If
    MapJobField wsMap, rr, "Applicant", NR_APPLICANT:                  rr = rr + 1
    ' Output Folder: canonical here for the inspector; a read-only mirror on the
    ' standard product (canonical on its Start Here) so the name isn't defined twice.
    If ProductIsInspector() Then
        MapJobField wsMap, rr, "Output Folder (optional)", NR_OUTFOLDER
    Else
        With wsMap.Range(wsMap.Cells(rr, MAP_JOB_LABEL_COL), wsMap.Cells(rr, MAP_JOB_LABEL_COL + 1))
            .Merge: .Value = "Output Folder (optional)": .Font.Bold = True
        End With
        With wsMap.Range(wsMap.Cells(rr, MAP_JOB_VALUE_COL), wsMap.Cells(rr, MAP_JOB_VALUE_LAST_COL))
            .Merge
            .Formula = "=IF(" & NR_OUTFOLDER & "="""",""(set on Start Here)""," & NR_OUTFOLDER & ")"
            .Font.Color = RGB(90, 90, 90): .Font.Italic = True
        End With
    End If
    rr = rr + 1

    ' Fetch Imagery source override (both products): blank = Esri World Imagery;
    ' the inspector can paste any other ArcGIS MapServer URL right here (per user
    ' request - the imagery source is swappable without touching the Sources
    ' sheet). Must be a MapServer: only those expose the /export operation.
    MapJobField wsMap, rr, "Imagery URL (optional)", NR_IMAGERYSVC
    rr = rr + 1

    ' ---- right-hand previews (formula-driven, update live as the user types) ----
    ' Stamp preview: a formula-built replica of the textbox stamped on every map
    ' page, driven by the FIRST Sites data row + the job boxes, so the inspector
    ' can see what each page's stamp will read BEFORE clicking Create Combined
    ' Map Pages PDF. Cols I:M, rows aligned with the job block; the "Job info"
    ' label row carries a matching caption on the right.
    With wsMap.Cells(MAP_JOB_FIRST_ROW - 1, 9)
        .Value = "Stamp preview (first Sites row):"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(47, 79, 79)
    End With
    With wsMap.Range(wsMap.Cells(MAP_JOB_FIRST_ROW, 9), wsMap.Cells(MAP_JOB_FIRST_ROW + 5, MAP_COLS_WIDE))
        .Merge
        .Formula = StampPreviewFormula()
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Color = RGB(0, 0, 0)
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .IndentLevel = 1
    End With

    ' FIRMette / file-name preview directly UNDER the stamp box (same JobFileStem
    ' every export uses; FirmettePreview is volatile so it tracks the job cells).
    With wsMap.Range(wsMap.Cells(MAP_JOB_FIRST_ROW + 6, 9), wsMap.Cells(MAP_JOB_FIRST_ROW + 6, MAP_COLS_WIDE))
        .Merge
        .Formula = "=""File name:   ""&FirmettePreview()"
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(90, 90, 90)
        .HorizontalAlignment = xlLeft
    End With

    ' WO/DI note (left, below the job block; kept clear of the right previews).
    With wsMap.Range(wsMap.Cells(MAP_JOB_FIRST_ROW + 7, MAP_JOB_LABEL_COL), wsMap.Cells(MAP_JOB_FIRST_ROW + 8, 8))
        .Merge
        .Value = "WO # and DI # here fill in blank WO #/DI # cells on the Sites tab (a value typed on the row wins)."
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(120, 120, 120)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
End Sub

' The stamp-preview cell formula: a live replica of BuildMapTextboxString driven
' by the FIRST Sites data row (row 2) and the job named ranges. Blank lines are
' skipped (each present line is prefixed with CHAR(10); the leading one is
' stripped with MID(...,2)). Kept as one self-contained formula - no helper
' cells - so the header band stays clean. Line order matches the VBA stamp:
' WO/DI, Applicant, Site, lat/lon, Category+Desc, Cost, Work.
Private Function StampPreviewFormula() As String
    Dim q As String, EMP As String
    q = Chr$(34)
    EMP = q & q                                  ' the "" empty-string literal

    ' WO/DI honour a row override then the Setup box, same as the real stamp.
    Dim effWO As String, effDI As String
    effWO = "IF(Sites!$A$2<>" & EMP & ",Sites!$A$2,JobWO)"
    effDI = "IF(Sites!$B$2<>" & EMP & ",Sites!$B$2,JobDI)"

    Dim ln1 As String, appL As String, siteL As String, latL As String
    Dim catL As String, costL As String, workL As String
    ln1 = "TRIM(IF(" & effWO & "<>" & EMP & "," & q & "WO #" & q & "&" & effWO & "," & EMP & ")&" & _
          q & " " & q & "&IF(" & effDI & "<>" & EMP & "," & q & "DI #" & q & "&" & effDI & "," & EMP & "))"
    appL = "JobApplicant"
    siteL = "IF(Sites!$C$2<>" & EMP & "," & q & "Site " & q & "&Sites!$C$2&" & q & " " & q & _
            "&Sites!$D$2,Sites!$D$2)"
    latL = "IF(AND(Sites!$E$2<>" & EMP & ",Sites!$F$2<>" & EMP & ")," & _
           "TEXT(Sites!$E$2," & q & "0.00000" & q & ")&" & q & ", " & q & _
           "&TEXT(Sites!$F$2," & q & "0.00000" & q & ")," & EMP & ")"
    catL = "IF(AND(Sites!$I$2<>" & EMP & ",Sites!$G$2<>" & EMP & ")," & _
           q & "Cat " & q & "&Sites!$I$2&" & q & ", " & q & "&Sites!$G$2," & _
           "IF(Sites!$I$2<>" & EMP & "," & q & "Cat " & q & "&Sites!$I$2," & _
           "IF(Sites!$G$2<>" & EMP & ",Sites!$G$2," & EMP & ")))"
    costL = "IF(Sites!$J$2<>" & EMP & "," & q & "Cost: " & q & "&Sites!$J$2," & EMP & ")"
    workL = "IF(Sites!$K$2<>" & EMP & "," & q & "Work: " & q & "&Sites!$K$2," & EMP & ")"

    Dim big As String
    big = PreviewNl(ln1) & "&" & PreviewNl(appL) & "&" & PreviewNl(siteL) & "&" & _
          PreviewNl(latL) & "&" & PreviewNl(catL) & "&" & PreviewNl(costL) & "&" & PreviewNl(workL)

    StampPreviewFormula = "=IF(LEN(" & big & ")=0," & q & "(fill in the job info + first Sites row)" & q & _
                          ",MID(" & big & ",2,32767))"
End Function

' IF(expr<>"", CHAR(10)&expr, "")  - one stamp line, blank-skipping.
Private Function PreviewNl(ByVal expr As String) As String
    Dim EMP As String
    EMP = Chr$(34) & Chr$(34)
    PreviewNl = "IF(" & expr & "<>" & EMP & ",CHAR(10)&" & expr & "," & EMP & ")"
End Function

' State dropdown validation for the Map Pages State cell (mirrors modBuild's).
Private Sub AddMapStateValidation(ByVal cell As Range)
    On Error Resume Next
    With cell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=STATE_LIST
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    On Error GoTo 0
End Sub

Private Sub HeaderLabel(ByVal wsMap As Worksheet, ByVal r As Long, ByVal c As Long, ByVal txt As String)
    With wsMap.Cells(r, c)
        .Value = txt
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(47, 79, 79)
    End With
End Sub

' One label + input cell in the header band. Never overwrites a value the user
' already typed (BuildWorkbook / Repair Layout re-runs this).
Private Sub MapJobField(ByVal wsMap As Worksheet, ByVal r As Long, _
        ByVal label As String, ByVal namedRange As String)
    With wsMap.Range(wsMap.Cells(r, MAP_JOB_LABEL_COL), wsMap.Cells(r, MAP_JOB_LABEL_COL + 1))
        .Merge
        .Value = label
        .Font.Bold = True
        .HorizontalAlignment = xlLeft
    End With
    With wsMap.Range(wsMap.Cells(r, MAP_JOB_VALUE_COL), wsMap.Cells(r, MAP_JOB_VALUE_LAST_COL))
        .Merge
        .Interior.Color = RGB(255, 255, 204)     ' same input yellow as Start Here
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlLeft
    End With
    AddNameForCell wsMap.Cells(r, MAP_JOB_VALUE_COL), namedRange
End Sub

' Everything actionable in the header band, all shapes (MAP_CTRL_PREFIX named).
' Idempotent - safe to call on every EnsureMapPagesSheet / PrepareMapPages.
' Redesigned 2026-07-15 per user direction ("one button; sideline the rest"):
'   - ONE hero button: "Create Map Pages PDF" (Prepare -> Fetch -> Export)
'   - "Download FIRMettes" beside it (the other primary deliverable)
'   - everything else - the individual steps, the re-run/re-stamp refreshers
'     and the manual Google Earth screenshot flow - lives under a collapsed
'     "Advanced options" toggle so the default view stays uncluttered.
' All shapes sit ABOVE the print area so none of it prints.
Private Sub AddMapPageControls(ByVal wsMap As Worksheet)
    Const GREEN As Long = 4563272          ' primary (matches the Start Here "Go")
    Const BLUE As Long = 12419407          ' Browse / secondary action on a field
    Dim shp As Shape, i As Long

    ' Idempotent: drop any existing control shapes first.
    For i = wsMap.Shapes.Count To 1 Step -1
        Set shp = wsMap.Shapes(i)
        If Left$(shp.Name, Len(MAP_CTRL_PREFIX)) = MAP_CTRL_PREFIX Then shp.Delete
    Next i

    Const GREY As Long = 7895160           ' "Exports & other tools" door

    ' Title.
    MapCtrlLabel wsMap, "Title", 8, 5, 300, 22, "Map Pages", 15, True, RGB(47, 79, 79)

    ' ---- the two primary deliverables, side by side ----
    ' Hero: one click does the whole combined map flow (Prepare -> Fetch -> Export).
    Dim hero As Shape
    Set hero = wsMap.Shapes.AddShape(msoShapeRoundedRectangle, 8, 32, 300, 42)
    hero.Name = MAP_CTRL_PREFIX & "CreatePdf"
    hero.Fill.ForeColor.RGB = GREEN
    hero.Line.Visible = msoFalse
    hero.Shadow.Visible = msoFalse
    With hero.TextFrame2.TextRange
        .Text = "Create Combined Map Pages PDF"
        .Font.Size = 13
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = vbWhite
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    hero.TextFrame2.VerticalAnchor = msoAnchorMiddle
    hero.OnAction = "CreateMapPagesPdf"

    ' Download FIRMettes right beside it (the other primary deliverable).
    Dim firm As Shape
    Set firm = wsMap.Shapes.AddShape(msoShapeRoundedRectangle, 320, 32, 220, 42)
    firm.Name = MAP_CTRL_PREFIX & "Firm"
    firm.Fill.ForeColor.RGB = GREEN
    firm.Line.Visible = msoFalse
    firm.Shadow.Visible = msoFalse
    With firm.TextFrame2.TextRange
        .Text = "Download FIRMettes"
        .Font.Size = 13
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = vbWhite
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    firm.TextFrame2.VerticalAnchor = msoAnchorMiddle
    firm.OnAction = "DownloadFirmettes"

    ' ---- Advanced options toggle (collapsed by default) ----
    ' Expansion state rides in the toggle shape's AlternativeText ("1" = open),
    ' so SetMapEditControlsVisible can restore the right state after an export.
    Set shp = wsMap.Shapes.AddShape(msoShapeRoundedRectangle, 8, 84, 160, 17)
    shp.Name = MAP_CTRL_PREFIX & "AdvToggle"
    shp.AlternativeText = "0"
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(150, 150, 150)
    shp.Line.Weight = 0.75
    shp.Shadow.Visible = msoFalse
    With shp.TextFrame2.TextRange
        .Text = "Advanced options  " & ChrW$(9656)          ' U+25B8 right triangle
        .Font.Size = 9
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.OnAction = "ToggleMapAdvanced"

    ' ---- Advanced content (every shape named MapCtrl_Adv*; hidden by default) ----
    ' Everything fits above y=204 (= the top of the "Job info" label row 13).
    ' Row 1: the three steps the hero chains, runnable one at a time, plus the
    ' ghost refresh/re-run actions.
    MapCtrlLabel wsMap, "Adv_StepsLabel", 8, 106, 250, 12, _
        "Run the steps one at a time:", 8, True, RGB(90, 90, 90)
    MapRibbonStep wsMap, "Adv_Prepare", 8, 120, 118, 18, BLUE, _
        "1. Prepare Pages", "PrepareMapPages", ""
    MapRibbonStep wsMap, "Adv_Fetch", 132, 120, 118, 18, BLUE, _
        "2. Fetch Imagery", "FetchMapImagery", ""
    MapRibbonStep wsMap, "Adv_Export", 256, 120, 118, 18, BLUE, _
        "3. Export PDF", "ExportCombinedMapPdf", ""
    MapGhostButton wsMap, "Adv_FetchRe", 388, 120, 138, 18, _
        ChrW$(8635) & " Re-run failed imagery", "ReRunFailedImagery"
    MapGhostButton wsMap, "Adv_Restamp", 532, 120, 108, 18, _
        ChrW$(8635) & " Re-stamp pages", "UpdateMapStamps"
    MapGhostButton wsMap, "Adv_FirmRe", 646, 120, 142, 18, _
        ChrW$(8635) & " Re-run failed FIRMettes", "ReRunFailedFirmettes"

    ' Row 2: an individual-PDFs export + the door to the Tools & Exports sheet.
    MapRibbonStep wsMap, "Adv_Individual", 8, 142, 236, 18, BLUE, _
        "Create Individual Map Pages PDFs", "CreateIndividualMapPagePdfs", ""
    If ProductIsInspector() Then
        MapRibbonStep wsMap, "Adv_Tools", 250, 142, 190, 18, GREY, _
            "Exports & other tools  " & ChrW$(8594), "GoToOtherTools", ""
    End If

    ' Row 3: the manual Google Earth screenshot flow.
    MapCtrlLabel wsMap, "Adv_ManualLabel", 8, 164, 300, 12, _
        "Manual alternative - Google Earth screenshots:", 8, True, RGB(90, 90, 90)
    MapRibbonStep wsMap, "Adv_KML", 8, 178, 118, 18, BLUE, _
        "Export to KML", "ExportSitesToKML", ""
    MapRibbonStep wsMap, "Adv_Insert", 132, 178, 118, 18, BLUE, _
        "Insert Images", "InsertMapImages", ""
    MapCtrlLabel wsMap, "Adv_ManualNote", 256, 174, 400, 22, _
        "KML opens in Google Earth Desktop - screenshot each site (Win+Shift+S), save as " & _
        "Site_1, Site_2..., then Insert Images. 'Re-stamp pages' refreshes stamps after job-info edits.", _
        8, False, RGB(110, 110, 110)

    ' ---- by the job boxes ----
    Dim reLeft As Double
    reLeft = wsMap.Cells(MAP_JOB_FIRST_ROW, MAP_JOB_VALUE_LAST_COL + 1).Left + 6

    ' Browse next to the Output Folder value - inspector only (Output Folder is
    ' canonical on MapPages there). The standard product browses on Start Here.
    If ProductIsInspector() Then
        ' Output Folder is the sixth job field: WO,DI,Disaster,State,Applicant,
        ' Output Folder = MAP_JOB_FIRST_ROW+5 on the inspector.
        Dim brTop As Double
        brTop = wsMap.Cells(MAP_JOB_FIRST_ROW + 5, 1).Top + 1
        Set shp = wsMap.Shapes.AddShape(msoShapeRoundedRectangle, reLeft, brTop, 52, 15)
        shp.Name = MAP_CTRL_PREFIX & "Browse"
        shp.Fill.ForeColor.RGB = BLUE
        shp.Line.Visible = msoFalse
        shp.Shadow.Visible = msoFalse
        With shp.TextFrame2.TextRange
            .Text = "Browse"
            .Font.Size = 9
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = vbWhite
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
        shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
        shp.OnAction = "SelectOutputFolder"
    End If

    ' Note beside the Imagery URL field (the row after Output Folder: +6 on the
    ' inspector, whose job block also carries State; +5 on the standard product).
    Dim imTop As Double
    imTop = wsMap.Cells(MAP_JOB_FIRST_ROW + IIf(ProductIsInspector(), 6, 5), 1).Top
    MapCtrlLabel wsMap, "ImgSvcNote", reLeft, imTop, 105, 30, _
        "Esri World Imagery by Default.", _
        8, False, RGB(120, 120, 120)

    ' Collapse the advanced shapes to the toggle's remembered state (fresh
    ' controls default to "0" = collapsed).
    ApplyAdvancedVisibility wsMap
End Sub

' Ghost button style: white fill, grey outline + text. Reads as a secondary
' refresh action, distinct from the green hero and the blue steps.
Private Sub MapGhostButton(ByVal wsMap As Worksheet, ByVal key As String, _
        ByVal leftPt As Double, ByVal topPt As Double, ByVal w As Double, ByVal h As Double, _
        ByVal caption As String, ByVal macroName As String)
    Dim shp As Shape
    Set shp = wsMap.Shapes.AddShape(msoShapeRoundedRectangle, leftPt, topPt, w, h)
    shp.Name = MAP_CTRL_PREFIX & key
    shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.Line.Visible = msoTrue
    shp.Line.ForeColor.RGB = RGB(150, 150, 150)
    shp.Line.Weight = 0.75
    shp.Shadow.Visible = msoFalse
    With shp.TextFrame2.TextRange
        .Text = caption
        .Font.Size = 8
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = RGB(90, 90, 90)
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
    shp.OnAction = macroName
End Sub

' Button: the "Advanced options" toggle - shows/hides every MapCtrl_Adv* shape.
Public Sub ToggleMapAdvanced()
    If Not SheetExists(SH_MAPPAGES) Then Exit Sub
    Dim wsMap As Worksheet, tog As Shape
    Set wsMap = ThisWorkbook.Worksheets(SH_MAPPAGES)
    On Error Resume Next
    Set tog = wsMap.Shapes(MAP_CTRL_PREFIX & "AdvToggle")
    On Error GoTo 0
    If tog Is Nothing Then Exit Sub
    tog.AlternativeText = IIf(tog.AlternativeText = "1", "0", "1")
    ApplyAdvancedVisibility wsMap
End Sub

Private Function AdvancedExpanded(ByVal wsMap As Worksheet) As Boolean
    On Error Resume Next
    AdvancedExpanded = (wsMap.Shapes(MAP_CTRL_PREFIX & "AdvToggle").AlternativeText = "1")
    On Error GoTo 0
End Function

' Show/hide the MapCtrl_Adv* shapes per the toggle's state and refresh the
' toggle caption's expand/collapse arrow.
Private Sub ApplyAdvancedVisibility(ByVal wsMap As Worksheet)
    Dim expanded As Boolean, shp As Shape
    Const ADV As String = "MapCtrl_Adv"
    expanded = AdvancedExpanded(wsMap)
    For Each shp In wsMap.Shapes
        If Left$(shp.Name, Len(ADV)) = ADV And shp.Name <> MAP_CTRL_PREFIX & "AdvToggle" Then
            shp.Visible = IIf(expanded, msoTrue, msoFalse)
        End If
    Next shp
    On Error Resume Next
    wsMap.Shapes(MAP_CTRL_PREFIX & "AdvToggle").TextFrame2.TextRange.Text = _
        "Advanced options  " & IIf(expanded, ChrW$(9662), ChrW$(9656))
    On Error GoTo 0
End Sub

' A non-interactive text label in the header band.
Private Sub MapCtrlLabel(ByVal wsMap As Worksheet, ByVal key As String, _
        ByVal leftPt As Double, ByVal topPt As Double, ByVal w As Double, ByVal h As Double, _
        ByVal txt As String, ByVal sz As Double, ByVal bold As Boolean, ByVal clr As Long)
    Dim lbl As Shape
    Set lbl = wsMap.Shapes.AddTextbox(msoTextOrientationHorizontal, leftPt, topPt, w, h)
    lbl.Name = MAP_CTRL_PREFIX & key
    lbl.Line.Visible = msoFalse
    lbl.Fill.Visible = msoFalse
    With lbl.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        With .TextRange
            .Text = txt
            .Font.Size = sz
            .Font.Bold = IIf(bold, msoTrue, msoFalse)
            .Font.Fill.ForeColor.RGB = clr
        End With
    End With
End Sub

' One workflow button + a tiny one-line caption below it. Button and caption are
' both MAP_CTRL_PREFIX named so SetMapEditControlsVisible can hide them.
Private Sub MapRibbonStep(ByVal wsMap As Worksheet, ByVal key As String, _
        ByVal leftPt As Double, ByVal topPt As Double, ByVal w As Double, ByVal h As Double, _
        ByVal fillColor As Long, ByVal caption As String, ByVal macroName As String, ByVal noteText As String)
    Dim btn As Shape
    Set btn = wsMap.Shapes.AddShape(msoShapeRoundedRectangle, leftPt, topPt, w, h)
    btn.Name = MAP_CTRL_PREFIX & key
    btn.Fill.ForeColor.RGB = fillColor
    btn.Line.Visible = msoFalse
    btn.Shadow.Visible = msoFalse
    With btn.TextFrame2.TextRange
        .Text = caption
        .Font.Size = 11
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = vbWhite
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    btn.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btn.OnAction = macroName

    If Len(noteText) = 0 Then Exit Sub
    ' 34pt tall so a two-line caption (e.g. the KML step) fits.
    MapCtrlLabel wsMap, key & "_Note", leftPt + 2, topPt + h + 2, w - 2, 34, noteText, 8, False, RGB(110, 110, 110)
End Sub

Private Sub ConfigureMapPageSetup(ByVal wsMap As Worksheet)
    With wsMap.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperLetter
        ' A real margin, not 0: see MAP_PRINT_MARGIN_PTS in modConstants - the
        ' content blocks are sized 756x576 to sit inside this frame, which is
        ' what makes the export exactly one PDF page per map page on ANY
        ' printer driver. Print 1:1 (no fit-to scaling - it ignores the manual
        ' page breaks and shrinks the content into loose whitespace).
        .LeftMargin = MAP_PRINT_MARGIN_PTS:  .RightMargin = MAP_PRINT_MARGIN_PTS
        .TopMargin = MAP_PRINT_MARGIN_PTS:   .BottomMargin = MAP_PRINT_MARGIN_PTS
        .HeaderMargin = 0: .FooterMargin = 0
        .CenterHorizontally = True
        .CenterVertically = True
        .Zoom = 100
        .PrintGridlines = False
        .PrintHeadings = False
    End With
    SizeMapColumns wsMap
End Sub

' Fit the 13 map columns to the map-block width (MAP_PAGE_WIDTH_PTS = the Letter
' width minus the print margins). ColumnWidth is in character units, not points,
' so we scale by the measured .Width (points) and converge; then guarantee the
' total never EXCEEDS the target (an over-width grid spills into a second
' horizontal page).
Private Sub SizeMapColumns(ByVal wsMap As Worksheet)
    Const TARGET As Double = MAP_PAGE_WIDTH_PTS
    Dim cc As Long, total As Double, factor As Double, iter As Long
    For cc = 1 To MAP_COLS_WIDE
        wsMap.Columns(cc).ColumnWidth = 10.5
    Next cc
    For iter = 1 To 4
        total = 0
        For cc = 1 To MAP_COLS_WIDE
            total = total + wsMap.Columns(cc).Width
        Next cc
        If total <= 0 Then Exit Sub
        If Abs(total - TARGET) < 0.5 Then Exit For
        factor = TARGET / total
        For cc = 1 To MAP_COLS_WIDE
            wsMap.Columns(cc).ColumnWidth = wsMap.Columns(cc).ColumnWidth * factor
        Next cc
    Next iter

    ' Trim any residual overage off the last column so total <= 792.
    total = 0
    For cc = 1 To MAP_COLS_WIDE
        total = total + wsMap.Columns(cc).Width
    Next cc
    If total > TARGET Then
        Dim lastW As Double
        lastW = wsMap.Columns(MAP_COLS_WIDE).Width
        If lastW > (total - TARGET) Then
            wsMap.Columns(MAP_COLS_WIDE).ColumnWidth = _
                wsMap.Columns(MAP_COLS_WIDE).ColumnWidth * ((lastW - (total - TARGET)) / lastW)
        End If
    End If
End Sub

' Lay out one page (4 merged rows x 13 cols) plus the WO/DI/applicant textbox.
' If wsSites is Nothing the page is blank (used by AddMapPage).
Private Sub CreateMapPage(ByVal wsMap As Worksheet, ByVal wsSites As Worksheet, _
        ByVal siteRow As Long, ByVal pageIdx As Long)
    Dim startRow As Long, rr As Long, rowH As Double, pageTopPts As Double
    Dim placeholderTxt As String, sNum As String, sNam As String
    Dim txtBox As Shape, txtContent As String, firstLineLen As Long

    rowH = MAP_PAGE_HEIGHT_PTS / MAP_ROWS_PER_PAGE
    ' Pages start BELOW the header band (which holds the job inputs).
    startRow = MAP_FIRST_PAGE_ROW + pageIdx * MAP_ROWS_PER_PAGE

    For rr = startRow To startRow + MAP_ROWS_PER_PAGE - 1
        wsMap.Rows(rr).RowHeight = rowH
    Next rr

    wsMap.Range(wsMap.Cells(startRow, 1), wsMap.Cells(startRow + MAP_ROWS_PER_PAGE - 1, MAP_COLS_WIDE)).Merge

    ' Faint placeholder = just which page this is. The old "Paste screenshot here
    ' (Place in Cell)" line was dropped per user request - Place-in-Cell paste
    ' isn't the real path (Insert Map Images / Select photo are), so the hint was
    ' misleading. Images cover this text once inserted.
    If wsSites Is Nothing Then
        placeholderTxt = "New Page"
    Else
        sNum = Trim$(CStr(wsSites.Cells(siteRow, COL_SITENO).Value))
        sNam = Trim$(CStr(wsSites.Cells(siteRow, COL_SITENAME).Value))
        If Len(sNum) > 0 Then
            placeholderTxt = "Site " & sNum & " " & sNam
        Else
            placeholderTxt = sNam
        End If
    End If
    With wsMap.Cells(startRow, 1)
        .Value = placeholderTxt
        .Font.Color = RGB(210, 210, 210)
        .Font.Size = 16
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
        .Placement = xlMove          ' never let row drift resize a shape (§9.8)
        ' Remember which Sites row this page came from, so UpdateMapStamps can
        ' re-derive the stamp after the job info changes. Blank pages get "0".
        .AlternativeText = CStr(IIf(wsSites Is Nothing, 0, siteRow))
        With .TextFrame2
            .WordWrap = msoTrue
            .MarginLeft = 5:  .MarginRight = 5
            .MarginTop = 3:    .MarginBottom = 3
            .AutoSize = msoAutoSizeNone
        End With
        With .TextFrame
            .Characters.Text = txtContent
            .Characters.Font.Name = "Segoe UI"
            .Characters.Font.Size = MAP_STAMP_FONT
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

    ' Per-page "Select photo" button in the top-right of the page area. Bound to
    ' this page index (not the site #), so it works even when the same Site #
    ' appears on several pages. Hidden during PDF export (SetMapEditControlsVisible).
    AddPickButton wsMap, pageIdx, pageTopPts

    ' Page break AFTER this page so the next CreateMapPage call lands on a new sheet page.
    On Error Resume Next
    wsMap.HPageBreaks.Add Before:=wsMap.Rows(startRow + MAP_ROWS_PER_PAGE)
    On Error GoTo 0
End Sub

' Small blue rounded-rectangle button, top-right of the page area, wired to
' modMapImage.PickImageForPage. Its name encodes the 1-based page number.
Private Sub AddPickButton(ByVal wsMap As Worksheet, ByVal pageIdx As Long, _
        ByVal pageTopPts As Double)
    Const BTN_W As Double = 92, BTN_H As Double = 20, PAD As Double = 6
    Dim btn As Shape
    Set btn = wsMap.Shapes.AddShape(msoShapeRoundedRectangle, _
        MAP_PAGE_WIDTH_PTS - BTN_W - PAD, pageTopPts + PAD, BTN_W, BTN_H)
    With btn
        .Name = MAP_PICKBTN_PREFIX & CStr(pageIdx + 1)
        .OnAction = "PickImageForPage"
        .Placement = xlMove          ' never let row drift resize a shape (§9.8)
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(0, 112, 192)
        End With
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = "Select photo"
            .Font.Size = 9
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With
End Sub

' Hide/show the on-sheet editing aids (per-page "Select photo" buttons and the
' off-grid "Insert Map Images" control + its note) - they're editing aids, not
' print content, and hiding the off-grid control also keeps it out of the used
' range at export time.
Public Sub SetMapEditControlsVisible(ByVal wsMap As Worksheet, ByVal vis As Boolean)
    Dim shp As Shape, vv As Long
    vv = IIf(vis, msoTrue, msoFalse)
    For Each shp In wsMap.Shapes
        If Left$(shp.Name, Len(MAP_PICKBTN_PREFIX)) = MAP_PICKBTN_PREFIX _
           Or Left$(shp.Name, Len(MAP_CTRL_PREFIX)) = MAP_CTRL_PREFIX Then
            shp.Visible = vv
        End If
    Next shp
    ' Restoring after an export must respect the Advanced section's collapsed
    ' state - otherwise every export would pop the advanced controls open.
    If vis Then ApplyAdvancedVisibility wsMap
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

    ' Build the stamp top-down, skipping the WO/DI line entirely when neither ID
    ' is set - and likewise the applicant line. Applicant used to be emitted
    ' unconditionally, so a blank one left an EMPTY first line in the textbox
    ' (and, since the bold run is "everything before the first vbLf", bolded
    ' nothing). That hit every RoadReviewer map page - the standard product has
    ' no Applicant field at all - and any inspector job that left it blank.
    ' Note WO/DI are NOT inspector-only: the Sites table carries them in cols
    ' A/B on both products (hidden in the standard one), and the row's own value
    ' wins over the Start Here field, so a standard-product user who unhides
    ' them and types an ID still gets it stamped.
    Dim costs As String, workComp As String
    costs = Trim$(CStr(wsSites.Cells(r, COL_COSTS).Value))
    workComp = Trim$(CStr(wsSites.Cells(r, COL_WORKCOMP).Value))

    If Len(woDiLine) > 0 Then BuildMapTextboxString = woDiLine & vbLf
    If Len(applicant) > 0 Then BuildMapTextboxString = BuildMapTextboxString & applicant & vbLf
    BuildMapTextboxString = BuildMapTextboxString & _
        siteLine & vbLf & _
        lat & ", " & lon
    If Len(catLine) > 0 Then BuildMapTextboxString = BuildMapTextboxString & vbLf & catLine
    ' Two optional money/progress lines, only emitted when populated.
    If Len(costs) > 0 Then BuildMapTextboxString = BuildMapTextboxString & vbLf & "Cost: " & costs
    If Len(workComp) > 0 Then BuildMapTextboxString = BuildMapTextboxString & vbLf & "Work: " & workComp
End Function

' The ONE-CLICK hero (user request 2026-07-15): Prepare Pages -> Fetch
' Imagery -> Export PDF as a single button, one summary message at the end.
' The individual steps stay available under "Advanced options" for partial
' re-runs and for the manual Google Earth screenshot flow.
Public Sub CreateMapPagesPdf()
    Dim pages As Long, ok As Long, failed As Long, pdfPath As String

    pages = PreparePagesCore()
    If pages < 1 Then Exit Sub          ' no rows / build error - already reported

    FetchImageryCore False, ok, failed

    pdfPath = ExportMapPdfCore()
    If Len(pdfPath) = 0 Then Exit Sub   ' export error - already reported

    If Not gHeadless Then
        Dim msg As String
        msg = "Map Pages PDF created." & vbCrLf & vbCrLf & _
              "Pages: " & pages & vbCrLf & _
              "Aerial imagery placed: " & ok & IIf(failed > 0, "   (failed: " & failed & ")", "") & _
              vbCrLf & vbCrLf & "Saved to:" & vbCrLf & pdfPath
        If failed > 0 Then msg = msg & vbCrLf & vbCrLf & _
            "Some imagery could not be downloaded (see the Map Status column on Sites). " & _
            "Fix the rows, then use Advanced options: 'Re-run failed imagery', then 'Export PDF'."
        MsgBox msg, IIf(failed = 0, vbInformation, vbExclamation), "Create Map Pages PDF"
    End If
End Sub

' Advanced-options step 3 (also runnable on its own).
Public Sub ExportCombinedMapPdf()
    Dim pdfPath As String
    pdfPath = ExportMapPdfCore()
    If Len(pdfPath) > 0 And Not gHeadless Then
        MsgBox "Combined map PDF exported:" & vbCrLf & pdfPath, vbInformation, "Export Map PDF"
    End If
End Sub

' Advanced option: one PDF PER SITE instead of a single combined file. Each is
' named exactly like that site's FIRMette but with "Location Map" in place of
' "FIRMette" (JobFileStem - <site name> - Location Map.pdf), so a site's map and
' FIRMette sit side by side in the folder. Uses the same direct writer as the
' combined export (one page each); the print pipeline is the per-page fallback.
Public Sub CreateIndividualMapPagePdfs()
    Dim wsMap As Worksheet, folder As String
    Dim nPages As Long, pageIdx As Long, ok As Long, failed As Long
    Dim siteName As String, fileName As String, fullPath As String, emsg As String
    Dim onePage(0 To 0) As Long

    If Not SheetExists(SH_MAPPAGES) Then
        If Not gHeadless Then MsgBox "No '" & SH_MAPPAGES & "' sheet yet. Click 'Create Combined Map Pages PDF' first.", _
            vbExclamation, "Individual Map PDFs"
        Exit Sub
    End If
    Set wsMap = ThisWorkbook.Worksheets(SH_MAPPAGES)
    ShowMapPages

    nPages = MapPageCount(wsMap)
    If nPages < 1 Then
        If Not gHeadless Then MsgBox "No map pages found. Click 'Create Combined Map Pages PDF' (or '1. Prepare Pages') first.", _
            vbExclamation, "Individual Map PDFs"
        Exit Sub
    End If

    folder = ResolveOutputFolder()
    If Not EnsureFolderExists(folder) Then
        If Not gHeadless Then MsgBox "Could not create the output folder:" & vbCrLf & folder, vbExclamation, "Individual Map PDFs"
        Exit Sub
    End If

    EnsureTextboxesOnTop wsMap
    NormalizeMapLayoutForPrint wsMap

    For pageIdx = 0 To nPages - 1
        siteName = MapPageSiteName(wsMap, pageIdx)
        fileName = CleanFileName(JobFileStem() & " - " & siteName & " Location Map.pdf")
        fullPath = folder & fileName
        SetStatus "Exporting map page " & (pageIdx + 1) & " of " & nPages & " - " & siteName
        DoEvents

        onePage(0) = pageIdx
        emsg = ""
        If BuildMapPdfForPages(wsMap, onePage, fullPath, emsg) Then
            ok = ok + 1
        ElseIf ExportOnePagePrint(wsMap, pageIdx, fullPath) Then
            ok = ok + 1
        Else
            failed = failed + 1
        End If
        DoEvents
    Next pageIdx
    ClearStatus
    If ok > 0 Then SurfaceFolder folder

    If Not gHeadless Then
        MsgBox "Individual map-page PDFs created." & vbCrLf & vbCrLf & _
            "Files written: " & ok & IIf(failed > 0, "   (failed: " & failed & ")", "") & vbCrLf & _
            "One PDF per site, named like the FIRMettes (""... Location Map.pdf"")." & vbCrLf & vbCrLf & _
            "Folder: " & folder, _
            IIf(failed = 0, vbInformation, vbExclamation), "Individual Map PDFs"
    End If
End Sub

' The site name for a map page (from the page stamp's remembered Sites row).
' Falls back to "Page N" for a blank/manual page so a file is still named.
Private Function MapPageSiteName(ByVal wsMap As Worksheet, ByVal pageIdx As Long) As String
    Dim siteRow As Long, nm As String
    On Error Resume Next
    siteRow = CLng(Val(wsMap.Shapes("Textbox_Page_" & CStr(pageIdx + 1)).AlternativeText))
    On Error GoTo 0
    If siteRow >= SITES_FIRST_DATA_ROW Then nm = Trim$(CStr(SitesSheet().Cells(siteRow, COL_SITENAME).Value))
    If Len(nm) = 0 Then nm = "Page " & CStr(pageIdx + 1)
    MapPageSiteName = nm
End Function

' Print-pipeline fallback for a SINGLE page (only when the direct writer is
' unavailable). Mirrors ExportMapPdfCore's fallback but with the print area
' restricted to the one page's rows. True on success.
Private Function ExportOnePagePrint(ByVal wsMap As Worksheet, ByVal pageIdx As Long, _
        ByVal fullPath As String) As Boolean
    Dim savedPrinter As String, borderless As Boolean, r1 As Long, r2 As Long, savedAlerts As Boolean
    On Error GoTo Fail
    SetMapEditControlsVisible wsMap, False
    On Error Resume Next
    savedPrinter = Application.ActivePrinter
    On Error GoTo Fail
    borderless = SwitchToPdfPrinter()
    ConfigureMapPageSetup wsMap
    NormalizeMapLayoutForPrint wsMap
    r1 = MAP_FIRST_PAGE_ROW + pageIdx * MAP_ROWS_PER_PAGE
    r2 = r1 + MAP_ROWS_PER_PAGE - 1
    savedAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    wsMap.PageSetup.PrintArea = wsMap.Range(wsMap.Cells(r1, 1), wsMap.Cells(r2, MAP_COLS_WIDE)).Address
    Application.DisplayAlerts = savedAlerts
    If Not borderless Then
        With wsMap.PageSetup
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
        End With
    End If
    wsMap.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fullPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
    RestorePrinter savedPrinter
    SetMapPrintArea wsMap
    SetMapEditControlsVisible wsMap, True
    ExportOnePagePrint = True
    Exit Function
Fail:
    RestorePrinter savedPrinter
    On Error Resume Next
    SetMapPrintArea wsMap
    SetMapEditControlsVisible wsMap, True
    On Error GoTo 0
End Function

' Exports the map pages to "<stem> - Location Map.pdf". Returns the full path,
' or "" when preconditions failed / the export errored non-headless.
Public Function ExportMapPdfCore() As String
    Dim wsMap As Worksheet, folder As String, fileName As String, fullPath As String

    If Not SheetExists(SH_MAPPAGES) Then
        If Not gHeadless Then MsgBox "No '" & SH_MAPPAGES & "' sheet. Click 'Create Map Pages PDF' first.", vbExclamation, "Export Map PDF"
        Exit Function
    End If
    Set wsMap = ThisWorkbook.Worksheets(SH_MAPPAGES)
    ShowMapPages

    folder = ResolveOutputFolder()
    If Not EnsureFolderExists(folder) Then
        If Not gHeadless Then MsgBox "Could not create the output folder:" & vbCrLf & folder, vbExclamation, "Export Map PDF"
        Exit Function
    End If
    fileName = CleanFileName(JobFileStem() & " - Location Map.pdf")
    fullPath = folder & fileName

    ' Bring text boxes to the front in case the inspector's pasted screenshot covered them.
    EnsureTextboxesOnTop wsMap
    ' Keep the on-screen geometry canonical (self-healing for drifted sheets;
    ' also what the print fallback below relies on).
    NormalizeMapLayoutForPrint wsMap

    ' PRIMARY: write the PDF directly from the page shapes + image files
    ' (modPdf). No printer driver involved, so the machine-specific
    ' print-render distortion (see modPdf's header) can never touch it.
    Dim directErr As String
    If BuildMapPdfDirect(wsMap, fullPath, directErr) Then
        SurfaceFolder folder
        ExportMapPdfCore = fullPath
        Exit Function
    End If
    TraceLine "Direct PDF export unavailable (" & directErr & ") - using the print-driver fallback"

    ' FALLBACK: the §9.8 print pipeline. Hide the on-sheet editing aids
    ' (per-page "Select photo" buttons + the off-grid controls) so they
    ' don't print.
    SetMapEditControlsVisible wsMap, False

    ' Export through "Microsoft Print to PDF": Excel paginates against the
    ' ACTIVE printer's usable area, which differs per driver (a Brother laser
    ' measured 749x552 usable; MS Print to PDF measures 769.5x576) - so pinning
    ' the export to the inbox MS PDF driver makes the pagination identical on
    ' every Windows machine. The 760x568 blocks are sized to ITS floor. Two
    ' hard-won rules: (1) switching the active printer RESETS the sheet's page
    ' setup (orientation flips back to portrait), so the setup must be
    ' re-applied AFTER the switch; (2) fit-to-page scaling is not a substitute
    ' - it ignores the manual page breaks and floats shrunken content in
    ' whitespace - so it's only the fallback when MS Print to PDF is absent.
    Dim savedPrinter As String, borderless As Boolean
    On Error Resume Next
    savedPrinter = Application.ActivePrinter
    On Error GoTo 0
    borderless = SwitchToPdfPrinter()
    On Error Resume Next
    ConfigureMapPageSetup wsMap
    ' Self-healing geometry (2026-07-15): re-assert row heights + manual page
    ' breaks from the constants and re-pin every printed shape to its page
    ' block. A sheet whose rows drifted (older-build layout, hand edits,
    ' OneDrive AutoSave persisting a half-migrated state) used to print the
    ' images stretched past the page blocks - the "screenshots outside the
    ' print area" bug. Now the export normalizes everything first.
    NormalizeMapLayoutForPrint wsMap
    SetMapPrintArea wsMap
    If Not borderless Then
        With wsMap.PageSetup
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = MapPageCount(wsMap)
        End With
    End If
    On Error GoTo 0

    On Error GoTo Fail
    wsMap.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fullPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
    On Error GoTo 0
    RestorePrinter savedPrinter
    SetMapEditControlsVisible wsMap, True
    SurfaceFolder folder
    ExportMapPdfCore = fullPath
    Exit Function
Fail:
    RestorePrinter savedPrinter
    SetMapEditControlsVisible wsMap, True
    ExportMapPdfCore = ""
    If gHeadless Then Err.Raise Err.Number, "ExportCombinedMapPdf", Err.Description Else _
        MsgBox "Export failed: " & Err.Description, vbCritical, "Export Map PDF"
End Function

' Re-assert the printed geometry from the current constants: header + page row
' heights, one manual page break after each page block, and every printed
' shape snapped back onto its block (modMapImage.SnapShapesToPages). Wrapped
' in On Error Resume Next - hardening must never block an export.
Private Sub NormalizeMapLayoutForPrint(ByVal wsMap As Worksheet)
    Dim nPages As Long, r As Long, pageIdx As Long, rowH As Double
    nPages = MapPageCount(wsMap)
    If nPages < 1 Then Exit Sub

    On Error Resume Next
    For r = 1 To MAP_HEADER_ROWS
        wsMap.Rows(r).RowHeight = MAP_HEADER_ROW_HEIGHT
    Next r
    rowH = MAP_PAGE_HEIGHT_PTS / MAP_ROWS_PER_PAGE
    For r = MAP_FIRST_PAGE_ROW To MAP_FIRST_PAGE_ROW + nPages * MAP_ROWS_PER_PAGE - 1
        wsMap.Rows(r).RowHeight = rowH
    Next r

    wsMap.ResetAllPageBreaks
    For pageIdx = 0 To nPages - 1
        wsMap.HPageBreaks.Add Before:=wsMap.Rows(MAP_FIRST_PAGE_ROW + (pageIdx + 1) * MAP_ROWS_PER_PAGE)
    Next pageIdx

    SnapShapesToPages wsMap, nPages
    On Error GoTo 0
End Sub

' Make "Microsoft Print to PDF" the active printer for the export. The reliable
' port comes from the registry (HKCU\...\Devices holds "winspool,Ne0X:" per
' printer - the exact suffix Excel's ActivePrinter wants); the Ne-port probe is
' only the fallback. NB: Excel rejects ActivePrinter changes when no workbook
' is open - always true here since we're exporting one.
Private Function SwitchToPdfPrinter() As Boolean
    Const PDF_PRINTER As String = "Microsoft Print to PDF"
    Dim port As String, i As Long
    On Error Resume Next
    port = CreateObject("WScript.Shell").RegRead( _
        "HKCU\Software\Microsoft\Windows NT\CurrentVersion\Devices\" & PDF_PRINTER)
    If InStr(port, ",") > 0 Then
        port = Mid$(port, InStr(port, ",") + 1)          ' "winspool,Ne00:" -> "Ne00:"
        Err.Clear
        Application.ActivePrinter = PDF_PRINTER & " on " & port
        If Err.Number = 0 Then SwitchToPdfPrinter = True: Exit Function
    End If
    For i = 0 To 31                                       ' fallback: probe the ports
        Err.Clear
        Application.ActivePrinter = PDF_PRINTER & " on Ne" & Format$(i, "00") & ":"
        If Err.Number = 0 Then SwitchToPdfPrinter = True: Exit Function
    Next i
    On Error GoTo 0
End Function

Private Sub RestorePrinter(ByVal savedPrinter As String)
    On Error Resume Next
    If Len(savedPrinter) > 0 Then Application.ActivePrinter = savedPrinter
    On Error GoTo 0
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
    SurfaceFolder ResolveOutputFolder()
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
            "No AGOL Webmap URL set. Paste your map's URL on the Start Here sheet " & _
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

' Open the state functional-class (NFC) layer in AGOL Map Viewer and drop
' ALL sites onto it at once, colored by federal-aid verdict. Reuses the
' colored KML (red = federal aid, green = non-federal, blue = review) and
' opens the same per-state NFC layer the "AGOL NFC Layer" column uses,
' centered on the first site so the map lands somewhere relevant. The user
' then drags the highlighted KML from Explorer onto the map.
Public Sub OpenSitesOnNfcLayer()
    Dim file As String, n As Long, dialogTitle As String
    dialogTitle = "Open Sites on NFC Layer"
    If Not WriteSitesKml(file, n, dialogTitle) Then Exit Sub

    ' Also drop a GeoJSON next to the KML: it's the better format for the AGOL
    ' Experience "Add Data" + Feature Info step-through (queryable feature layer
    ' + real verdict fields). Failure here is non-fatal - the KML still works.
    Dim gjFile As String, gjN As Long
    WriteSitesGeoJson gjFile, gjN, dialogTitle

    Dim url As String
    url = NfcLayerUrlForFirstSite()
    If Len(url) = 0 Then url = BuildUrl(NfcLayerTemplate(), 0, 0)   ' no sites w/ coords - open layer anyway

    If Not gHeadless Then
        On Error Resume Next
        ThisWorkbook.FollowHyperlink Address:=url, NewWindow:=False
        On Error GoTo 0
        Dim q As String: q = Chr$(34)
        Shell "explorer.exe /select," & q & file & q, vbNormalFocus
        MsgBox "Exported " & n & " point(s) to:" & vbCrLf & file & vbCrLf & _
            IIf(Len(gjFile) > 0, "and GeoJSON: " & gjFile & vbCrLf, "") & vbCrLf & _
            "The state NFC functional-class layer should now be open in ArcGIS Map Viewer." & vbCrLf & _
            "Drag the highlighted KML from Explorer onto the map to add all your sites," & vbCrLf & _
            "colored red (federal aid) / green (non-federal aid) / blue (review)." & vbCrLf & vbCrLf & _
            "For an ArcGIS Experience with a click-through review, add the .geojson " & _
            "instead (Add Data widget) - see docs/agol-review-app.md.", _
            vbInformation, dialogTitle
    End If
End Sub

' The per-state AGOL NFC-layer URL template (matches the AGOL NFC Layer
' column - modBuild.SetNfcAgolFormula / modExport.NfcAgolUrlForRow).
Private Function NfcLayerTemplate() As String
    Select Case BareStateCode(SetupValue(NR_STATE))
        Case "IN": NfcLayerTemplate = URL_NFC_MAPVIEW_IN
        Case "WI": NfcLayerTemplate = URL_NFC_MAPVIEW_WI_LOCAL   ' local-first, matches the primary link column (PR #37)
        Case "MN": NfcLayerTemplate = URL_NFC_MAPVIEW_MN
        Case "IL": NfcLayerTemplate = URL_NFC_MAPVIEW_IL
        Case "OH": NfcLayerTemplate = URL_NFC_MAPVIEW_OH
        Case "MI", "": NfcLayerTemplate = URL_NFC_MAPVIEW
        Case Else: NfcLayerTemplate = URL_FEMAVIEW
    End Select
End Function

Private Function NfcLayerUrlForFirstSite() As String
    Dim ws As Worksheet, last As Long, r As Long
    Set ws = SitesSheet()
    last = SitesLastRow()
    For r = SITES_FIRST_DATA_ROW To last
        If HasValidCoords(ws, r) Then
            NfcLayerUrlForFirstSite = BuildUrl(NfcLayerTemplate(), ws.Cells(r, COL_LAT).Value, ws.Cells(r, COL_LON).Value)
            Exit Function
        End If
    Next r
End Function

' Shared KML builder: writes the file to the resolved output folder, sets
' filePath + placemarkCount on success, returns False on any failure (with a
' MsgBox already shown if gHeadless is False).
' Param deliberately not named featureCount - that would collide
' (case-insensitively) with modHttp's FeatureCount() function; see the
' NfcWired/nfcWired compile-error note in modClassify.bas for why that's a
' real trap, not just a style nit.
Private Function WriteSitesKml(ByRef filePath As String, ByRef placemarkCount As Long, _
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
        "<name>" & XmlEscape(ProductTitle() & " Sites") & "</name>" & vbCrLf & _
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
    file = folder & CleanFileName(JobFileStem() & " - Sites.kml")

    If Not WriteTextFile(file, kml) Then
        If Not gHeadless Then MsgBox "Could not write the KML file.", vbExclamation, dialogTitle
        Exit Function
    End If

    filePath = file
    placemarkCount = n
    WriteSitesKml = True
End Function

' Walk the Sites table twice:
'   pass 1 — collect every unique Category value (trimmed, case-preserved
'            but case-insensitive grouping) into the order it first
'            appears, so the folder list is predictable
'   pass 2 — for each Category bucket, emit a <Folder> with all its rows
' Rows with a blank Category fall into a "(no category)" folder so they
' aren't lost. Updates placemarkCount with the number of placemarks emitted.
Private Function BuildCategoryFolders(ByVal ws As Worksheet, ByVal last As Long, _
        ByRef placemarkCount As Long) As String
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
            placemarkCount = placemarkCount + 1
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

' ---- GeoJSON export (ArcGIS Online Add Data / Experience Builder path) -----
' Prefer GeoJSON over KML for the AGOL path: a GeoJSON added at runtime becomes
' a fully queryable client-side FEATURE layer, so Experience Builder's Feature
' Info "1 of N" navigation, field-based selection, and symbolize-by-field all
' work. An added KML becomes a limited "KML layer" that widgets often can't step
' through. Each site's classifier results ride along as flat PROPERTIES (real
' AGOL fields), including a Verdict bucket + VerdictColor hex so the layer can be
' styled red/green/blue by field. KML export stays for Google Earth.
' See docs/agol-review-app.md.
Public Sub ExportSitesToGeoJson()
    Dim file As String, n As Long, dialogTitle As String
    dialogTitle = "Export GeoJSON"
    If Not WriteSitesGeoJson(file, n, dialogTitle) Then Exit Sub
    SurfaceFolder ResolveOutputFolder()
    If Not gHeadless Then
        MsgBox "Exported " & n & " point(s) to:" & vbCrLf & file & vbCrLf & vbCrLf & _
            "In your ArcGIS Online Experience, use the Add Data widget to load this " & _
            ".geojson, then step through the sites with the Feature Info widget.", _
            vbInformation, dialogTitle
    End If
End Sub

' Shared GeoJSON builder: mirrors WriteSitesKml. Writes a FeatureCollection of
' every row with valid coordinates to the resolved output folder. Sets
' filePath + featureCount on success; returns False (with a MsgBox unless
' headless) on failure.
Private Function WriteSitesGeoJson(ByRef filePath As String, ByRef featureCount As Long, _
        ByVal dialogTitle As String) As Boolean
    Dim ws As Worksheet, last As Long, r As Long, feats As String, n As Long
    Set ws = SitesSheet()
    last = SitesLastRow()
    If last < SITES_FIRST_DATA_ROW Then
        If Not gHeadless Then MsgBox "No site rows to export.", vbInformation, dialogTitle
        Exit Function
    End If

    For r = SITES_FIRST_DATA_ROW To last
        If HasValidCoords(ws, r) Then
            If n > 0 Then feats = feats & "," & vbCrLf
            feats = feats & SiteFeatureJson(ws, r)
            n = n + 1
        End If
    Next r

    If n = 0 Then
        If Not gHeadless Then MsgBox "No rows have valid coordinates to export.", vbInformation, dialogTitle
        Exit Function
    End If

    Dim js As String
    js = "{""type"":""FeatureCollection"",""features"":[" & vbCrLf & feats & vbCrLf & "]}"

    Dim folder As String, file As String
    folder = ResolveOutputFolder()
    If Not EnsureFolderExists(folder) Then
        If Not gHeadless Then MsgBox "Could not create the output folder:" & vbCrLf & folder, vbExclamation, dialogTitle
        Exit Function
    End If
    file = folder & CleanFileName(JobFileStem() & " - Sites.geojson")

    If Not WriteTextFile(file, js) Then
        If Not gHeadless Then MsgBox "Could not write the GeoJSON file.", vbExclamation, dialogTitle
        Exit Function
    End If

    filePath = file
    featureCount = n
    WriteSitesGeoJson = True
End Function

' One GeoJSON Feature for a Sites row. Geometry is [lon, lat] per RFC 7946
' (GeoJSON is always WGS84, so AGOL ingests it without reprojection). Optional
' properties are omitted when blank so the field list stays clean.
Private Function SiteFeatureJson(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim lat As String, lon As String, status As String, nm As String, props As String
    lat = InvariantNum(ws.Cells(r, COL_LAT).Value)
    lon = InvariantNum(ws.Cells(r, COL_LON).Value)
    status = Trim$(CStr(ws.Cells(r, COL_ELIGIBILITY).Value))
    nm = Trim$(CStr(ws.Cells(r, COL_SITENAME).Value))
    If Len(nm) = 0 Then nm = "Site row " & r

    props = """Name"":""" & JsonEsc(nm) & """"     ' Name always present (first, no leading comma)
    props = props & JProp("SiteNo", CStr(ws.Cells(r, COL_SITENO).Value))
    props = props & JProp("WO", CStr(ws.Cells(r, COL_WO).Value))
    props = props & JProp("DI", CStr(ws.Cells(r, COL_DI).Value))
    props = props & JProp("FedAidStatus", status)
    props = props & JProp("Verdict", VerdictBucketLabel(status))
    props = props & JProp("VerdictColor", VerdictColorHex(status))
    props = props & JProp("FHWAClass", CStr(ws.Cells(r, COL_CLASS).Value))
    props = props & JProp("UrbanRural", CStr(ws.Cells(r, COL_URBANRURAL).Value))
    props = props & JProp("ACUBName", CStr(ws.Cells(r, COL_ACUBNAME).Value))
    props = props & JProp("RoadName", CStr(ws.Cells(r, COL_ROADNAME).Value))
    props = props & JProp("StreetName", CStr(ws.Cells(r, COL_STREET).Value))
    props = props & JProp("ReviewNote", CStr(ws.Cells(r, COL_REVIEWNOTE).Value))
    props = props & JProp("Category", CStr(ws.Cells(r, COL_CATEGORY).Value))
    props = props & JProp("Description", BuildDescBlock(ws, r, ""))
    ' Numeric lat/lon too - handy in an attribute table / label.
    props = props & ",""Latitude"":" & lat & ",""Longitude"":" & lon

    SiteFeatureJson = "{""type"":""Feature"",""properties"":{" & props & _
        "},""geometry"":{""type"":""Point"",""coordinates"":[" & lon & "," & lat & "]}}"
End Function

' Emit a JSON  ,"key":"value"  pair (leading comma), or "" when the value is
' blank. Callers seed the object with the always-present Name pair first, so
' every optional pair can safely lead with a comma.
Private Function JProp(ByVal key As String, ByVal raw As String) As String
    Dim v As String: v = Trim$(raw)
    If Len(v) = 0 Then Exit Function
    JProp = ",""" & key & """:""" & JsonEsc(v) & """"
End Function

' Verdict bucket + display color, reusing the same status->bucket mapping the
' KML pins use (PinStyleId). Colors match the web prototype / PDF palette.
Private Function VerdictBucketLabel(ByVal status As String) As String
    Select Case PinStyleId(status)
        Case "fedAid": VerdictBucketLabel = "Federal aid"
        Case "nonFedAid": VerdictBucketLabel = "Non-federal aid"
        Case Else: VerdictBucketLabel = "Review"
    End Select
End Function

Private Function VerdictColorHex(ByVal status As String) As String
    Select Case PinStyleId(status)
        Case "fedAid": VerdictColorHex = "#c0392b"       ' red
        Case "nonFedAid": VerdictColorHex = "#27ae60"    ' green
        Case Else: VerdictColorHex = "#2980b9"           ' blue (review / unclassified)
    End Select
End Function

' JSON string escape. Escapes the mandatory control/quote/backslash chars and,
' to keep the file pure-ASCII (WriteTextFile writes ANSI), any byte >126 as a
' \uXXXX sequence - which is valid UTF-8-safe JSON.
Private Function JsonEsc(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case """": out = out & "\"""
            Case "\": out = out & "\\"
            Case vbCr: out = out & "\r"
            Case vbLf: out = out & "\n"
            Case vbTab: out = out & "\t"
            Case Else
                code = AscW(ch) And &HFFFF&
                If code < 32 Or code > 126 Then
                    out = out & "\u" & Right$("000" & LCase$(Hex$(code)), 4)
                Else
                    out = out & ch
                End If
        End Select
    Next i
    JsonEsc = out
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
