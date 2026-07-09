Attribute VB_Name = "modMapImage"
Option Explicit

' RoadReviewer - Workflow 3 add-on: Insert Map Images.
'
' Fills the MapPages placeholders ("Paste screenshot here") automatically
' instead of by hand.
'
' WHY THE IMAGES COME FROM DISK, NOT FROM VBA
' -------------------------------------------
' The authoritative per-site figure (Esri street basemap + ACUB urban polygon
' + the state's functional-class polylines, each drawn with its OWN published
' drawingInfo.renderer) is a *vector composite*. Only MDOT (MI) exposes a
' MapServer/export operation; Indiana, Wisconsin and the nationwide ACUB layer
' are AGOL-hosted Query-only feature services with no export and no /legend
' (confirmed live 2026-07-03 - see web/index.html header comment). VBA has no
' canvas and no image compositor, so it cannot reproduce that figure. Porting
' it would give a Michigan-only raster path and silently drop IN/WI/ACUB.
'
' So: web/index.html renders the exact same figure it already draws for the
' PDF report, exports it as PNG (renderCombinedCanvas -> canvas.toBlob), and
' this module places those PNGs into the MapPages layout. One renderer, one
' symbology, no drift between the web tool and the workbook.
'
' EXPECTED FILE NAMES (either form, checked in this order):
'   Site_<SiteNo>.png     e.g. Site_12.png   (preferred - survives reordering)
'   Page_<n>.png          e.g. Page_3.png    (1-based page index fallback)
' Also accepts .jpg / .jpeg for hand-captured screenshots.
'
' Network: none. This module only touches the local filesystem.

Private Const IMG_SUBFOLDER As String = "maps\"
Private Const SHAPE_PREFIX As String = "MapImage_Page_"
Private Const IMG_INSET_PTS As Double = 4       ' hairline gap inside the page area

' ---- entry points ---------------------------------------------------------

' Button: "Insert Map Images" on the MapPages sheet (or the 3. Maps sheet).
Public Sub InsertMapImages()
    On Error GoTo Fail

    If Not SheetExists(SH_MAPPAGES) Then
        If Not gHeadless Then MsgBox "No '" & SH_MAPPAGES & "' sheet. Click 'Prepare Map Pages' first.", _
            vbExclamation, "Insert Map Images"
        Exit Sub
    End If

    Dim folder As String
    folder = ResolveMapImageFolder()
    If Len(folder) = 0 Then Exit Sub                 ' user cancelled the picker

    If Not FolderHasImages(folder) Then
        If Not gHeadless Then MsgBox _
            "No map images found in:" & vbCrLf & folder & vbCrLf & vbCrLf & _
            "In the web tool (index.html), click 'Download Map PNGs', then point " & _
            "this at the folder the browser saved them to." & vbCrLf & vbCrLf & _
            "Expected names: Site_<SiteNo>.png  or  Page_<n>.png", _
            vbInformation, "Insert Map Images"
        Exit Sub
    End If

    Dim wsMap As Worksheet, wsSites As Worksheet
    Set wsMap = ThisWorkbook.Worksheets(SH_MAPPAGES)
    Set wsSites = SitesSheet()

    ' Re-run safe: drop any pictures we previously inserted, keep textboxes.
    RemoveMapImages

    Dim siteRows As Collection
    Set siteRows = MapPageSiteRows(wsSites)

    Dim pageIdx As Long, placed As Long, missing As Long
    Dim imgPath As String, siteNo As String

    Application.ScreenUpdating = False
    For pageIdx = 0 To PageCount(wsMap) - 1
        siteNo = ""
        If pageIdx < siteRows.Count Then
            siteNo = Trim$(CStr(wsSites.Cells(CLng(siteRows(pageIdx + 1)), COL_SITENO).Value))
        End If

        imgPath = FindImageForPage(folder, siteNo, pageIdx + 1)
        If Len(imgPath) > 0 Then
            PlaceImageOnPage wsMap, pageIdx, imgPath
            placed = placed + 1
            SetStatus "Placing map image " & placed & "..."
            DoEvents
        Else
            missing = missing + 1
        End If
    Next pageIdx
    Application.ScreenUpdating = True
    ClearStatus

    If Not gHeadless Then
        Dim msg As String
        msg = placed & " map image(s) placed."
        If missing > 0 Then msg = msg & vbCrLf & missing & " page(s) had no matching file and kept their placeholder."
        MsgBox msg, vbInformation, "Insert Map Images"
    End If
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    ClearStatus
    If gHeadless Then
        Err.Raise Err.Number, "InsertMapImages", Err.Description
    Else
        MsgBox "Insert Map Images failed:" & vbCrLf & Err.Description, vbExclamation, "Insert Map Images"
    End If
End Sub

' Button: "Remove Map Images" - restores the paste-a-screenshot placeholders.
Public Sub RemoveMapImages()
    If Not SheetExists(SH_MAPPAGES) Then Exit Sub

    Dim wsMap As Worksheet, shp As Shape, i As Long
    Set wsMap = ThisWorkbook.Worksheets(SH_MAPPAGES)

    For i = wsMap.Shapes.Count To 1 Step -1
        Set shp = wsMap.Shapes(i)
        If Left$(shp.Name, Len(SHAPE_PREFIX)) = SHAPE_PREFIX Then shp.Delete
    Next i
End Sub

' ---- folder resolution ----------------------------------------------------

' Default: <output folder>\maps\ if it exists, else prompt. The picker is
' skipped entirely in headless verification runs.
Private Function ResolveMapImageFolder() As String
    Dim guess As String
    guess = ResolveOutputFolder() & IMG_SUBFOLDER
    If FolderHasImages(guess) Then
        ResolveMapImageFolder = guess
        Exit Function
    End If

    If gHeadless Then
        ResolveMapImageFolder = guess       ' caller reports "no images found"
        Exit Function
    End If

    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Choose the folder containing the exported map PNGs"
    fd.InitialFileName = ResolveOutputFolder()
    If fd.Show <> -1 Then Exit Function      ' cancelled -> ""

    Dim chosen As String
    chosen = fd.SelectedItems(1)
    If Right$(chosen, 1) <> "\" Then chosen = chosen & "\"
    ResolveMapImageFolder = chosen
End Function

Private Function FolderHasImages(ByVal folder As String) As Boolean
    If Len(folder) = 0 Then Exit Function
    If Len(Dir$(folder & "*.png")) > 0 Then FolderHasImages = True: Exit Function
    If Len(Dir$(folder & "*.jpg")) > 0 Then FolderHasImages = True: Exit Function
    If Len(Dir$(folder & "*.jpeg")) > 0 Then FolderHasImages = True
End Function

' Site_<n>.png preferred, Page_<n>.png fallback; png -> jpg -> jpeg.
Private Function FindImageForPage(ByVal folder As String, ByVal siteNo As String, _
                                  ByVal pageNum As Long) As String
    Dim stems As Variant, exts As Variant, s As Variant, e As Variant, p As String

    If Len(siteNo) > 0 Then
        stems = Array("Site_" & siteNo, "Page_" & pageNum)
    Else
        stems = Array("Page_" & pageNum)
    End If
    exts = Array(".png", ".jpg", ".jpeg")

    For Each s In stems
        For Each e In exts
            p = folder & s & e
            If Len(Dir$(p)) > 0 Then FindImageForPage = p: Exit Function
        Next e
    Next s
End Function

' ---- page geometry --------------------------------------------------------

' CreateMapPage lays out MAP_ROWS_PER_PAGE rows per page starting at
' pageIdx * MAP_ROWS_PER_PAGE + 1, merged across MAP_COLS_WIDE columns,
' with total height MAP_PAGE_HEIGHT_PTS. Pages are counted from the
' textboxes CreateMapPage stamps ("Textbox_Page_<n>").
Private Function PageCount(ByVal wsMap As Worksheet) As Long
    Dim shp As Shape, n As Long
    For Each shp In wsMap.Shapes
        If Left$(shp.Name, Len("Textbox_Page_")) = "Textbox_Page_" Then n = n + 1
    Next shp
    PageCount = n
End Function

Private Function PageTopPts(ByVal wsMap As Worksheet, ByVal pageIdx As Long) As Double
    Dim startRow As Long, rr As Long, t As Double
    startRow = pageIdx * MAP_ROWS_PER_PAGE + 1
    For rr = 1 To startRow - 1
        t = t + wsMap.Rows(rr).Height
    Next rr
    PageTopPts = t
End Function

Private Function PageWidthPts(ByVal wsMap As Worksheet) As Double
    Dim cc As Long, w As Double
    For cc = 1 To MAP_COLS_WIDE
        w = w + wsMap.Columns(cc).Width
    Next cc
    PageWidthPts = w
End Function

' ---- placement ------------------------------------------------------------

' Insert the picture, scale it to fit the page area preserving aspect ratio,
' center it, push it behind the WO/DI textbox, and clear the placeholder text
' from the merged cell underneath.
Private Sub PlaceImageOnPage(ByVal wsMap As Worksheet, ByVal pageIdx As Long, _
                             ByVal imgPath As String)
    Dim areaTop As Double, areaLeft As Double, areaW As Double, areaH As Double
    areaTop = PageTopPts(wsMap, pageIdx) + IMG_INSET_PTS
    areaLeft = IMG_INSET_PTS
    areaW = PageWidthPts(wsMap) - IMG_INSET_PTS * 2
    areaH = MAP_PAGE_HEIGHT_PTS - IMG_INSET_PTS * 2

    Dim shp As Shape
    ' LinkToFile:=False, SaveWithDocument:=True -> the workbook stays portable
    ' after the PNG folder is deleted.
    Set shp = wsMap.Shapes.AddPicture(Filename:=imgPath, _
        LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=areaLeft, Top:=areaTop, Width:=-1, Height:=-1)

    With shp
        .Name = SHAPE_PREFIX & CStr(pageIdx + 1)
        .LockAspectRatio = msoTrue

        ' Fit-inside: scale by the tighter of the two ratios.
        Dim scaleF As Double
        If .Width <= 0 Or .Height <= 0 Then Exit Sub
        scaleF = areaW / .Width
        If (areaH / .Height) < scaleF Then scaleF = areaH / .Height

        .Width = .Width * scaleF
        .Height = .Height * scaleF

        .Left = areaLeft + (areaW - .Width) / 2
        .Top = areaTop + (areaH - .Height) / 2

        .Placement = xlMoveAndSize
        .ZOrder msoSendToBack          ' keep the WO/DI textbox readable on top
    End With

    ClearPlaceholderText wsMap, pageIdx
End Sub

' The merged cell carries grey italic "Paste screenshot here" prompt text.
' Once a real image is in place that prompt would print over/behind it.
Private Sub ClearPlaceholderText(ByVal wsMap As Worksheet, ByVal pageIdx As Long)
    Dim startRow As Long
    startRow = pageIdx * MAP_ROWS_PER_PAGE + 1
    ' The page area is a MERGED cell. ClearContents on a single member cell
    ' raises "We can't do that to a merged cell" (and hangs under headless COM
    ' with a just-placed picture over it) - clear the whole MergeArea instead.
    ' Purely cosmetic (the picture already covers the text), so never let it
    ' break the run.
    On Error Resume Next
    wsMap.Cells(startRow, 1).MergeArea.ClearContents
    On Error GoTo 0
End Sub

' ---- site row order -------------------------------------------------------

' PrepareMapPages walks the Sites table top-to-bottom, one page per row that
' has valid coordinates. Mirror that order so page N lines up with site N.
Private Function MapPageSiteRows(ByVal wsSites As Worksheet) As Collection
    Dim col As New Collection, r As Long, lastRow As Long
    lastRow = wsSites.Cells(wsSites.Rows.Count, COL_LAT).End(xlUp).Row

    For r = SITES_FIRST_DATA_ROW To lastRow
        If HasValidCoords(wsSites, r) Then col.Add r
    Next r

    Set MapPageSiteRows = col
End Function
