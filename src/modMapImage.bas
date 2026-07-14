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
' HOW A FILE IS MATCHED TO A PAGE (resolution order, per page):
'   1. Page_<n>       e.g. Page_3.png    - explicit, always unique, wins.
'   2. Site_<SiteNo>  e.g. Site_12.png   - explicit, but ONLY when that Site #
'                                           appears on a single page. If the same
'                                           Site # is on several pages (same site,
'                                           different WO/DI), Site_ is ambiguous
'                                           and is skipped for those pages.
'   3. File order     - any page still unmatched is filled from the remaining
'                       images oldest->newest: natural filename sort (so
'                       "Screenshot (2)" precedes "(10)"), modified-time as the
'                       tiebreaker. A file already used by 1/2 is never reused.
' Also accepts .jpg / .jpeg for hand-captured screenshots.
'
' Per-page manual override: each map page carries a "Select photo" button
' (modMaps.AddPickButton -> PickImageForPage) that places a chosen file on that
' exact page, immune to the duplicate-Site# ambiguity because it's bound to the
' page position, not the name.
'
' Network: none. This module only touches the local filesystem.

Private Const IMG_SUBFOLDER As String = "maps\"
Private Const SHAPE_PREFIX As String = "MapImage_Page_"
Private Const IMG_INSET_PTS As Double = 0       ' 0 = image covers the full page area, no gap

' One image file discovered in the folder, for the oldest->newest auto-fill.
Private Type ImgFile
    Name As String
    Path As String
    Modt As Date
End Type

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

    Dim nPages As Long
    nPages = PageCount(wsMap)
    If nPages < 1 Then
        If Not gHeadless Then MsgBox "No map pages found. Click 'Prepare Map Pages' first.", _
            vbExclamation, "Insert Map Images"
        Exit Sub
    End If

    ' Site # for each page, in page order, + which site numbers are duplicated.
    Dim siteRows As Collection
    Set siteRows = MapPageSiteRows(wsSites)
    Dim pageSite() As String, pageIdx As Long
    ReDim pageSite(0 To nPages - 1)
    For pageIdx = 0 To nPages - 1
        If pageIdx < siteRows.Count Then
            pageSite(pageIdx) = Trim$(CStr(wsSites.Cells(CLng(siteRows(pageIdx + 1)), COL_SITENO).Value))
        End If
    Next pageIdx

    ' Ordered image list (oldest->newest: natural filename, modified-time tiebreak).
    Dim imgs() As ImgFile, nImgs As Long
    CollectImages folder, imgs, nImgs

    Dim consumed() As Boolean, pageImg() As String
    ReDim pageImg(0 To nPages - 1)
    If nImgs > 0 Then ReDim consumed(0 To nImgs - 1)

    ' Pass 1 - explicit Page_<n> (always) and Site_<n> (only if that site is unique).
    For pageIdx = 0 To nPages - 1
        Dim p As String
        p = ExplicitOverride(folder, pageSite(pageIdx), pageIdx + 1, _
                             SiteIsDuplicated(pageSite, pageIdx))
        If Len(p) > 0 Then
            pageImg(pageIdx) = p
            MarkConsumed imgs, nImgs, consumed, p
        End If
    Next pageIdx

    ' Pass 2 - fill the rest from remaining images in order.
    Dim k As Long
    k = 0
    For pageIdx = 0 To nPages - 1
        If Len(pageImg(pageIdx)) = 0 Then
            Do While k < nImgs
                If Not consumed(k) Then Exit Do
                k = k + 1
            Loop
            If k < nImgs Then
                pageImg(pageIdx) = imgs(k).Path
                consumed(k) = True
                k = k + 1
            End If
        End If
    Next pageIdx

    ' Place.
    Dim placed As Long, missing As Long
    Application.ScreenUpdating = False
    For pageIdx = 0 To nPages - 1
        If Len(pageImg(pageIdx)) > 0 Then
            PlaceImageOnPage wsMap, pageIdx, pageImg(pageIdx)
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
        If missing > 0 Then msg = msg & vbCrLf & missing & " page(s) had no image and kept their placeholder."
        If AnyDuplicateSite(pageSite) Then msg = msg & vbCrLf & vbCrLf & _
            "Note: a Site # appears on more than one page, so 'Site_<n>' naming " & _
            "was skipped for it. Use 'Page_<n>' names, file order, or a page's " & _
            "'Select photo' button to place those."
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

' Wired to each page's "Select photo" button (modMaps.AddPickButton). Reads the
' page index from the calling shape's name, prompts for one image file, and
' places it on that exact page - replacing whatever image was there.
Public Sub PickImageForPage()
    On Error GoTo Fail
    If gHeadless Then Exit Sub                       ' interactive-only

    Dim callerName As String, pageIdx As Long
    callerName = CStr(Application.Caller)
    If Left$(callerName, Len(MAP_PICKBTN_PREFIX)) <> MAP_PICKBTN_PREFIX Then Exit Sub
    pageIdx = CLng(Mid$(callerName, Len(MAP_PICKBTN_PREFIX) + 1)) - 1
    If pageIdx < 0 Then Exit Sub

    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Choose the image for this page"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Images", "*.png; *.jpg; *.jpeg"
    fd.InitialFileName = ResolveOutputFolder()
    If fd.Show <> -1 Then Exit Sub                   ' cancelled

    Dim wsMap As Worksheet
    Set wsMap = ThisWorkbook.Worksheets(SH_MAPPAGES)
    RemoveImageOnPage wsMap, pageIdx
    PlaceImageOnPage wsMap, pageIdx, fd.SelectedItems(1)
    Exit Sub
Fail:
    MsgBox "Could not place the image:" & vbCrLf & Err.Description, vbExclamation, "Select Photo"
End Sub

' Delete only the placed picture on one page (leaves the textbox + pick button).
Private Sub RemoveImageOnPage(ByVal wsMap As Worksheet, ByVal pageIdx As Long)
    Dim target As String
    target = SHAPE_PREFIX & CStr(pageIdx + 1)
    Dim shp As Shape, i As Long
    For i = wsMap.Shapes.Count To 1 Step -1
        Set shp = wsMap.Shapes(i)
        If shp.Name = target Then shp.Delete
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

' Explicit name match for one page: Page_<n> always; Site_<siteNo> only when
' that Site # is unique across the map pages. Returns "" if neither exists.
Private Function ExplicitOverride(ByVal folder As String, ByVal siteNo As String, _
                                  ByVal pageNum As Long, ByVal siteIsDup As Boolean) As String
    Dim p As String
    If Len(siteNo) > 0 And Not siteIsDup Then
        p = FirstExisting(folder, "Site_" & siteNo)
        If Len(p) > 0 Then ExplicitOverride = p: Exit Function
    End If
    ExplicitOverride = FirstExisting(folder, "Page_" & pageNum)
End Function

' First of stem.png / stem.jpg / stem.jpeg that exists in the folder.
Private Function FirstExisting(ByVal folder As String, ByVal stem As String) As String
    Dim exts As Variant, e As Variant, p As String
    exts = Array(".png", ".jpg", ".jpeg")
    For Each e In exts
        p = folder & stem & e
        If Len(Dir$(p)) > 0 Then FirstExisting = p: Exit Function
    Next e
End Function

' True if the Site # at pageSite(idx) is non-blank and appears on another page too.
Private Function SiteIsDuplicated(ByRef pageSite() As String, ByVal idx As Long) As Boolean
    Dim me_ As String, i As Long
    me_ = pageSite(idx)
    If Len(me_) = 0 Then Exit Function
    For i = LBound(pageSite) To UBound(pageSite)
        If i <> idx Then
            If StrComp(pageSite(i), me_, vbTextCompare) = 0 Then SiteIsDuplicated = True: Exit Function
        End If
    Next i
End Function

' True if any non-blank Site # is shared by two or more pages.
Private Function AnyDuplicateSite(ByRef pageSite() As String) As Boolean
    Dim i As Long
    For i = LBound(pageSite) To UBound(pageSite)
        If SiteIsDuplicated(pageSite, i) Then AnyDuplicateSite = True: Exit Function
    Next i
End Function

' All png/jpg/jpeg in the folder, sorted oldest->newest: natural filename order
' (numeric chunks compared as numbers), file modified-time as the tiebreaker.
Private Sub CollectImages(ByVal folder As String, ByRef imgs() As ImgFile, ByRef n As Long)
    Dim pats As Variant, pat As Variant, nm As String
    ReDim imgs(0 To 63)
    n = 0
    pats = Array("*.png", "*.jpg", "*.jpeg")
    Dim pi As Long
    For pi = LBound(pats) To UBound(pats)
        pat = pats(pi)
        nm = Dir$(folder & pat)
        Do While Len(nm) > 0
            If Not NameAlreadyListed(imgs, n, nm) Then
                If n > UBound(imgs) Then ReDim Preserve imgs(0 To UBound(imgs) + 64)
                imgs(n).Name = nm
                imgs(n).Path = folder & nm
                imgs(n).Modt = FileDateTime(folder & nm)
                n = n + 1
            End If
            nm = Dir$()
        Loop
    Next pi
    If n = 0 Then Exit Sub

    ' Insertion sort (small n): natural filename, then modified-time.
    Dim i As Long, j As Long, tmp As ImgFile
    For i = 1 To n - 1
        tmp = imgs(i)
        j = i - 1
        Do While j >= 0
            If ImgLess(tmp, imgs(j)) Then
                imgs(j + 1) = imgs(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        imgs(j + 1) = tmp
    Next i
End Sub

Private Function NameAlreadyListed(ByRef imgs() As ImgFile, ByVal n As Long, ByVal nm As String) As Boolean
    Dim i As Long
    For i = 0 To n - 1
        If StrComp(imgs(i).Name, nm, vbTextCompare) = 0 Then NameAlreadyListed = True: Exit Function
    Next i
End Function

' a < b : natural filename order, modified-time as tiebreaker.
Private Function ImgLess(ByRef a As ImgFile, ByRef b As ImgFile) As Boolean
    Dim c As Long
    c = NaturalCompare(a.Name, b.Name)
    If c <> 0 Then ImgLess = (c < 0): Exit Function
    ImgLess = (a.Modt < b.Modt)
End Function

' Mark the folder image whose path matches p (case-insensitive) as consumed, so
' the file-order pass never reuses a file an explicit name already claimed.
Private Sub MarkConsumed(ByRef imgs() As ImgFile, ByVal n As Long, _
                         ByRef consumed() As Boolean, ByVal p As String)
    Dim i As Long
    For i = 0 To n - 1
        If StrComp(imgs(i).Path, p, vbTextCompare) = 0 Then consumed(i) = True: Exit Sub
    Next i
End Sub

' Case-insensitive, numeric-aware string compare: -1 / 0 / 1.
Private Function NaturalCompare(ByVal a As String, ByVal b As String) As Long
    a = LCase$(a): b = LCase$(b)
    Dim i As Long, j As Long, la As Long, lb As Long
    i = 1: j = 1: la = Len(a): lb = Len(b)
    Do While i <= la And j <= lb
        Dim ca As String, cb As String
        ca = Mid$(a, i, 1): cb = Mid$(b, j, 1)
        If IsDigit(ca) And IsDigit(cb) Then
            Dim na As String, nb As String
            na = ""
            Do While i <= la
                If Not IsDigit(Mid$(a, i, 1)) Then Exit Do
                na = na & Mid$(a, i, 1)
                i = i + 1
            Loop
            nb = ""
            Do While j <= lb
                If Not IsDigit(Mid$(b, j, 1)) Then Exit Do
                nb = nb & Mid$(b, j, 1)
                j = j + 1
            Loop
            Dim va As Double, vb As Double
            va = Val(na): vb = Val(nb)
            If va < vb Then NaturalCompare = -1: Exit Function
            If va > vb Then NaturalCompare = 1: Exit Function
        Else
            If ca < cb Then NaturalCompare = -1: Exit Function
            If ca > cb Then NaturalCompare = 1: Exit Function
            i = i + 1: j = j + 1
        End If
    Loop
    If i > la And j > lb Then
        NaturalCompare = 0
    ElseIf i > la Then
        NaturalCompare = -1
    Else
        NaturalCompare = 1
    End If
End Function

Private Function IsDigit(ByVal ch As String) As Boolean
    IsDigit = (ch >= "0" And ch <= "9")
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

' Insert the picture, CROP it to the page area's aspect ratio (no distortion),
' then scale the cropped remainder to exactly fill the area edge-to-edge,
' push it behind the WO/DI textbox, and clear the placeholder text from the
' merged cell underneath.
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
        If .Width <= 0 Or .Height <= 0 Or areaW <= 0 Or areaH <= 0 Then Exit Sub

        ' "Cover" crop: trim the picture (in its own native point-space, which
        ' PictureFormat.Crop* always uses regardless of later resizing) down
        ' to the page area's aspect ratio, cutting evenly off both sides of
        ' whichever axis has surplus, then scale the trimmed remainder up to
        ' fill the area exactly. No stretching/distortion, no letterbox gaps.
        Dim natW As Double, natH As Double, targetAspect As Double, nativeAspect As Double
        natW = .Width
        natH = .Height
        targetAspect = areaW / areaH
        nativeAspect = natW / natH

        With .PictureFormat
            If nativeAspect > targetAspect Then
                Dim cropW As Double
                cropW = natW - natH * targetAspect
                .CropLeft = cropW / 2
                .CropRight = cropW / 2
                .CropTop = 0
                .CropBottom = 0
            Else
                Dim cropH As Double
                cropH = natH - natW / targetAspect
                .CropTop = cropH / 2
                .CropBottom = cropH / 2
                .CropLeft = 0
                .CropRight = 0
            End If
        End With

        .LockAspectRatio = msoFalse
        .Width = areaW
        .Height = areaH
        .Left = areaLeft
        .Top = areaTop

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
