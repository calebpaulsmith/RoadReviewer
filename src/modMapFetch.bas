Attribute VB_Name = "modMapFetch"
Option Explicit

' RoadReviewer - Map Pages: Fetch Imagery (the auto path, PR #35).
'
' For every map page that came from a Sites row, download an aerial image
' centered on the site from Esri's World Imagery export endpoint (one
' anonymous GET returns a rendered PNG for a Web-Mercator bbox - confirmed
' live 2026-07-14) and place it through the same crop-to-cover pipeline the
' manual Insert Map Images flow uses. This replaces Google-Earth
' screenshotting as the DEFAULT map-page flow; the KML + screenshots +
' Insert Images path stays available as the manual alternative for
' inspectors who prefer Google Earth's imagery or oblique angles.
'
' Because the site is always the exact center of the fetched bbox, a small
' printed pin shape is dropped at the geometric center of each page, plus
' the Esri-required attribution line in the bottom-left corner. Both use
' their own shape-name prefixes (MAP_PIN_PREFIX / MAP_ATTR_PREFIX) so
' SetMapEditControlsVisible never hides them - they must print.
'
' Per-page status is written to the Sites row's Map Status column with the
' shared STATUS_FAILED_PREFIX convention, so Re-run Failed Imagery retries
' only the rows that failed - same model as the FIRMette batch.
'
' Known risk: reachability of services.arcgisonline.com from a hardened
' FEMA laptop is untested (fine from the dev machine). If it's blocked,
' the endpoint is overridable via the Sources sheet's Service URLs table
' (key WORLD_IMAGERY), and the manual screenshot flow still works.

' Downloaded PNGs also land in <output folder>\maps\ under the same names
' Insert Map Images looks for (Site_<n> / Page_<n>), so the auto and manual
' flows interoperate on the same folder.
Private Const IMG_SUBFOLDER As String = "maps\"

' ---- entry points ---------------------------------------------------------

Public Sub FetchMapImagery()
    Dim ok As Long, failed As Long
    If Not FetchImageryCore(False, ok, failed) Then Exit Sub
    ShowFetchSummary ok, failed
End Sub

Public Sub ReRunFailedImagery()
    Dim ok As Long, failed As Long
    If Not FetchImageryCore(True, ok, failed) Then Exit Sub
    ShowFetchSummary ok, failed
End Sub

Private Sub ShowFetchSummary(ByVal ok As Long, ByVal failed As Long)
    If gHeadless Then Exit Sub
    MsgBox "Imagery fetch complete." & vbCrLf & _
        "Placed: " & ok & vbCrLf & _
        "Failed: " & failed & vbCrLf & vbCrLf & _
        "PNG copies: " & ResolveOutputFolder() & IMG_SUBFOLDER & vbCrLf & vbCrLf & _
        "Each image is centered on the site (the yellow pin marks it). " & _
        "Prefer Google Earth instead? Use the manual alternative under Advanced options.", _
        IIf(failed = 0, vbInformation, vbExclamation), "Fetch Imagery"
End Sub

' ---- batch driver ----------------------------------------------------------

' Downloads + places imagery for every (or just-failed) site page. Returns
' True when a run actually happened; okOut/failedOut carry the counts. The
' final summary MsgBox lives in the wrappers so CreateMapPagesPdf can chain
' this silently and report once at the end.
Public Function FetchImageryCore(ByVal onlyFailed As Boolean, _
        ByRef okOut As Long, ByRef failedOut As Long) As Boolean
    Dim wsMap As Worksheet, wsSites As Worksheet
    Dim nPages As Long, pageIdx As Long
    Dim total As Long, processed As Long, ok As Long, failed As Long

    okOut = 0
    failedOut = 0
    If Not SheetExists(SH_MAPPAGES) Then
        If Not gHeadless Then MsgBox "No '" & SH_MAPPAGES & "' sheet. Click 'Create Map Pages PDF' first.", _
            vbExclamation, "Fetch Imagery"
        Exit Function
    End If
    Set wsMap = ThisWorkbook.Worksheets(SH_MAPPAGES)
    Set wsSites = SitesSheet()

    nPages = MapPageCount(wsMap)
    If nPages < 1 Then
        If Not gHeadless Then MsgBox "No map pages found. Click 'Create Map Pages PDF' first.", _
            vbExclamation, "Fetch Imagery"
        Exit Function
    End If
    ' Reveal only once there's real work: the compile check no-op-runs this sub
    ' against the COMMITTED workbook, and an unconditional ShowMapPages up top
    ' unhid the standard product's Map Pages there (OneDrive AutoSave then
    ' persisted it - the §7d trap, caught by verify-skeleton).
    ShowMapPages

    Dim folder As String, mapsFolder As String
    folder = ResolveOutputFolder()
    mapsFolder = folder & IMG_SUBFOLDER
    ' Best effort - if the maps\ copy folder can't be created the run still
    ' works off %TEMP%; only the on-disk PNG copies are skipped.
    EnsureFolderExists mapsFolder

    For pageIdx = 0 To nPages - 1
        If ShouldFetchPage(wsMap, wsSites, pageIdx, onlyFailed) Then total = total + 1
    Next pageIdx
    If total = 0 Then
        If Not gHeadless Then MsgBox IIf(onlyFailed, "No failed imagery rows to re-run.", _
            "No site pages to fetch imagery for."), vbInformation, "Fetch Imagery"
        Exit Function
    End If

    ' Stay on Map Pages with redraw ON: each page is one network round trip,
    ' so the inspector watches the images appear page by page (same live-
    ' progress reasoning as the classify/FIRMette loops, §7a increment 5).
    wsMap.Activate
    For pageIdx = 0 To nPages - 1
        If Not ShouldFetchPage(wsMap, wsSites, pageIdx, onlyFailed) Then GoTo NextPage
        processed = processed + 1

        Dim siteRow As Long, siteName As String, msg As String
        siteRow = PageSiteRow(wsMap, pageIdx)
        siteName = Trim$(CStr(wsSites.Cells(siteRow, COL_SITENAME).Value))
        If Len(siteName) = 0 Then siteName = "row" & siteRow
        SetStatus "Fetching imagery " & processed & " of " & total & " - " & siteName
        DoEvents

        msg = ""
        If FetchOnePage(wsMap, wsSites, pageIdx, siteRow, mapsFolder, msg) Then
            wsSites.Cells(siteRow, COL_MAPSTATUS).Value = "Imagery placed: " & msg
            ok = ok + 1
        Else
            wsSites.Cells(siteRow, COL_MAPSTATUS).Value = STATUS_FAILED_PREFIX & "imagery: " & Left$(msg, 220)
            failed = failed + 1
        End If
        DoEvents
NextPage:
    Next pageIdx
    ClearStatus
    If ok > 0 Then SurfaceFolder folder

    okOut = ok
    failedOut = failed
    FetchImageryCore = True
End Function

' A page is fetchable when it references a Sites row (blank AddMapPage pages
' carry "0" and are skipped). Coordinate validity is deliberately NOT checked
' here: a row whose coords went bad after Prepare must still be PROCESSED so
' it gets a visible Failed status instead of being silently skipped.
Private Function ShouldFetchPage(ByVal wsMap As Worksheet, ByVal wsSites As Worksheet, _
        ByVal pageIdx As Long, ByVal onlyFailed As Boolean) As Boolean
    Dim siteRow As Long
    siteRow = PageSiteRow(wsMap, pageIdx)
    If siteRow < SITES_FIRST_DATA_ROW Then Exit Function
    If onlyFailed Then
        ShouldFetchPage = (InStr(1, CStr(wsSites.Cells(siteRow, COL_MAPSTATUS).Value), _
            STATUS_FAILED_PREFIX, vbTextCompare) = 1)
    Else
        ShouldFetchPage = True
    End If
End Function

' The Sites row a page came from - CreateMapPage bakes it into the page
' stamp's AlternativeText (same lookup UpdateMapStamps uses). 0 = blank page.
Private Function PageSiteRow(ByVal wsMap As Worksheet, ByVal pageIdx As Long) As Long
    On Error Resume Next
    PageSiteRow = CLng(Val(wsMap.Shapes("Textbox_Page_" & CStr(pageIdx + 1)).AlternativeText))
    On Error GoTo 0
End Function

' ---- one page: download -> place -> pin -> attribution ---------------------

' On success returns True with resultMsg = the placed file's name; on failure
' returns False with resultMsg = the error.
Private Function FetchOnePage(ByVal wsMap As Worksheet, ByVal wsSites As Worksheet, _
        ByVal pageIdx As Long, ByVal siteRow As Long, ByVal mapsFolder As String, _
        ByRef resultMsg As String) As Boolean
    If Not HasValidCoords(wsSites, siteRow) Then
        resultMsg = "invalid coordinates on Sites row " & siteRow
        Exit Function
    End If

    Dim lat As Double, lon As Double
    lat = CDbl(wsSites.Cells(siteRow, COL_LAT).Value)
    lon = CDbl(wsSites.Cells(siteRow, COL_LON).Value)

    Dim url As String, tempPath As String, errMsg As String
    url = ImageryExportUrl(lat, lon)
    tempPath = Environ$("TEMP") & "\rr_imagery_page_" & CStr(pageIdx + 1) & ".png"
    If Not HttpDownloadBinary(url, tempPath, "image/png,image/*,*/*", "image", errMsg) Then
        resultMsg = errMsg
        Exit Function
    End If

    ' Copy into <output folder>\maps\ under the Insert-Map-Images naming so the
    ' manual flow (and a re-run of it) sees the same files. Best effort only.
    Dim stem As String, destName As String
    stem = Trim$(CStr(wsSites.Cells(siteRow, COL_SITENO).Value))
    If Len(stem) > 0 Then
        destName = "Site_" & CleanFileName(stem) & ".png"
    Else
        destName = "Page_" & CStr(pageIdx + 1) & ".png"
    End If
    On Error Resume Next
    FileCopy tempPath, mapsFolder & destName
    On Error GoTo 0

    On Error GoTo Fail
    ' Same pipeline as the manual flow: RemoveImageOnPage also clears any
    ' previous pin/attribution for the page, so a re-fetch never stacks shapes.
    RemoveImageOnPage wsMap, pageIdx
    PlaceImageOnPage wsMap, pageIdx, tempPath
    ' The %TEMP% download gets overwritten by later runs; point the shape's
    ' remembered source (used by the direct PDF export) at the durable copy.
    If Len(Dir$(mapsFolder & destName)) > 0 Then TagImageSource wsMap, pageIdx, mapsFolder & destName
    AddSitePin wsMap, pageIdx
    AddAttribution wsMap, pageIdx
    On Error GoTo 0

    resultMsg = destName
    FetchOnePage = True
    Exit Function
Fail:
    resultMsg = "placing image: " & Err.Description
End Function

' Imagery export URL for a frame centered on (lat,lon): half-width
' MAP_IMG_HALFWIDTH_M, half-height scaled to the page-block aspect so the
' image fills the 760x568 block with zero cropping. All numbers written
' with invariant decimal points (Str$ always uses "."), so a comma-decimal
' locale can't corrupt the bbox.
Private Function ImageryExportUrl(ByVal lat As Double, ByVal lon As Double) As String
    Dim x As Double, y As Double, halfW As Double, halfH As Double
    x = WebMercX(lon)
    y = WebMercY(lat)
    halfW = MAP_IMG_HALFWIDTH_M
    halfH = halfW * MAP_PAGE_HEIGHT_PTS / MAP_PAGE_WIDTH_PTS
    ImageryExportUrl = ImageryServiceBase() & "/export?bbox=" & _
        InvariantD(x - halfW) & "," & InvariantD(y - halfH) & "," & _
        InvariantD(x + halfW) & "," & InvariantD(y + halfH) & _
        "&bboxSR=3857&imageSR=3857&size=" & MAP_IMG_PX_W & "," & MAP_IMG_PX_H & _
        "&format=png&transparent=false&f=image"
End Function

' Which ArcGIS MapServer to fetch from. The Map Pages "Imagery URL" cell wins
' when filled in (any ArcGIS MapServer works - only MapServers expose /export;
' a Query-only FeatureServer fails with a clear per-row error). Blank falls
' back to Esri World Imagery via the Svc_WORLD_IMAGERY override / default.
' Pasted URLs are normalized: query string, trailing slashes and a trailing
' /export are stripped, so copying the URL out of a browser's address bar
' after testing an export works as-is.
Private Function ImageryServiceBase() As String
    Dim v As String, p As Long
    v = Trim$(SetupValue(NR_IMAGERYSVC))
    If Len(v) = 0 Then v = ServiceUrl("WORLD_IMAGERY")
    p = InStr(v, "?")
    If p > 0 Then v = Left$(v, p - 1)
    Do While Right$(v, 1) = "/"
        v = Left$(v, Len(v) - 1)
    Loop
    If LCase$(Right$(v, 7)) = "/export" Then v = Left$(v, Len(v) - 7)
    ImageryServiceBase = v
End Function

' Attribution line for the fetched pages: the Esri credit for the default
' World Imagery service, or "Imagery: <host>" when the user re-pointed the
' fetch at another ArcGIS server.
Private Function AttributionText() As String
    If Len(Trim$(SetupValue(NR_IMAGERYSVC))) = 0 Then
        AttributionText = MAP_IMG_ATTRIBUTION
    Else
        AttributionText = "Imagery: " & UrlHost(ImageryServiceBase())
    End If
End Function

Private Function UrlHost(ByVal url As String) As String
    Dim s As String, p As Long
    s = url
    p = InStr(s, "://")
    If p > 0 Then s = Mid$(s, p + 3)
    p = InStr(s, "/")
    If p > 0 Then s = Left$(s, p - 1)
    UrlHost = s
End Function

' WGS84 lon/lat -> Web Mercator (EPSG:3857) meters.
Private Function WebMercX(ByVal lon As Double) As Double
    WebMercX = lon * 20037508.34 / 180#
End Function

Private Function WebMercY(ByVal lat As Double) As Double
    Const PI As Double = 3.14159265358979
    WebMercY = Log(Tan((90# + lat) * PI / 360#)) / (PI / 180#) * 20037508.34 / 180#
End Function

' Invariant-decimal string for a double ("." decimal separator, no thousands).
Private Function InvariantD(ByVal v As Double) As String
    InvariantD = Trim$(Str$(v))
End Function

' ---- printed overlays -------------------------------------------------------

' Google Earth-style yellow pushpin whose needle TIP sits exactly on the site
' (= the geometric center of the page block; the fetched image is bbox-centered
' on the site). Built from grouped shapes - a grey needle, a yellow ball head
' with a darker outline, and a small highlight glint - because no image asset
' ships with the workbook. Replaced the old red dot per user direction
' (2026-07-15). The group's bottom-center IS the site point; SnapShapesToPages
' relies on that convention when it re-pins the group at export time. Added
' AFTER the picture, which PlaceImageOnPage sent to back, so it draws on top.
Private Sub AddSitePin(ByVal wsMap As Worksheet, ByVal pageIdx As Long)
    Const HEAD_D As Double = 15        ' pushpin head diameter
    Const NEEDLE_W As Double = 4.5     ' needle width at the head
    Const NEEDLE_H As Double = 13      ' needle length (head -> tip)
    Const OVERLAP As Double = 2        ' head overlaps the needle top

    Dim cx As Double, cy As Double
    cx = PageWidthPts(wsMap) / 2#
    cy = PageTopPts(wsMap, pageIdx) + MAP_PAGE_HEIGHT_PTS / 2#

    Dim needle As Shape, head As Shape, glint As Shape, grp As Shape

    ' Needle: downward triangle, tip exactly on the site.
    Set needle = wsMap.Shapes.AddShape(msoShapeIsoscelesTriangle, _
        cx - NEEDLE_W / 2#, cy - NEEDLE_H, NEEDLE_W, NEEDLE_H)
    needle.Rotation = 180
    needle.Fill.ForeColor.RGB = RGB(110, 110, 110)
    needle.Line.Visible = msoTrue
    needle.Line.ForeColor.RGB = RGB(70, 70, 70)
    needle.Line.Weight = 0.5
    needle.Shadow.Visible = msoFalse

    ' Head: the yellow ball.
    Set head = wsMap.Shapes.AddShape(msoShapeOval, _
        cx - HEAD_D / 2#, cy - NEEDLE_H - HEAD_D + OVERLAP, HEAD_D, HEAD_D)
    head.Fill.ForeColor.RGB = RGB(255, 204, 0)
    head.Line.Visible = msoTrue
    head.Line.ForeColor.RGB = RGB(150, 110, 0)
    head.Line.Weight = 0.75
    head.Shadow.Visible = msoFalse

    ' Highlight glint, upper-left of the head (the Google Earth look).
    Set glint = wsMap.Shapes.AddShape(msoShapeOval, _
        head.Left + HEAD_D * 0.22, head.Top + HEAD_D * 0.18, HEAD_D * 0.3, HEAD_D * 0.3)
    glint.Fill.ForeColor.RGB = RGB(255, 244, 180)
    glint.Line.Visible = msoFalse
    glint.Shadow.Visible = msoFalse

    Set grp = wsMap.Shapes.Range(Array(needle.Name, head.Name, glint.Name)).Group
    With grp
        .Name = MAP_PIN_PREFIX & CStr(pageIdx + 1)
        .Placement = xlMove          ' never let row drift resize the pin (§9.8)
        .ZOrder msoBringToFront
    End With
End Sub

' Esri-required attribution, bottom-left of the page block, small enough to
' stay unobtrusive on the printed page. A textbox, so EnsureTextboxesOnTop
' keeps it above the picture at export time.
Private Sub AddAttribution(ByVal wsMap As Worksheet, ByVal pageIdx As Long)
    Const ATTR_W As Double = 220, ATTR_H As Double = 12
    Dim shp As Shape, t As Double
    t = PageTopPts(wsMap, pageIdx) + MAP_PAGE_HEIGHT_PTS - ATTR_H - 2
    Set shp = wsMap.Shapes.AddTextbox(msoTextOrientationHorizontal, 2, t, ATTR_W, ATTR_H)
    With shp
        .Name = MAP_ATTR_PREFIX & CStr(pageIdx + 1)
        .Placement = xlMove          ' never let row drift resize a shape (§9.8)
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
            .Transparency = 0.35
        End With
        .Line.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            With .TextRange
                .Text = AttributionText()
                .Font.Size = 7
                .Font.Fill.ForeColor.RGB = RGB(60, 60, 60)
            End With
        End With
    End With
End Sub
