Attribute VB_Name = "modPdf"
Option Explicit

' RoadReviewer - direct PDF writer for the combined Location Map (2026-07-16).
'
' WHY THIS EXISTS - THE PRINT-RENDER DISTORTION (worse than §9.8)
' ---------------------------------------------------------------
' ExportAsFixedFormat renders through Excel's print pipeline, and on real
' machines (confirmed on the dev laptop with BOTH the Brother driver and
' Microsoft Print to PDF, interactive AND headless, cells AND shapes, even a
' bare 100x100pt square on a fresh workbook) that pipeline can draw the page
' VERTICALLY STRETCHED ~5-7% while paginating with nominal sizes. The result
' the user caught: every map image printed ~760x606pt instead of 760x568,
' bleeding past the page block ("screenshots outside the print area"). No
' arrangement of shapes, rows, margins, printers or crops fixes a renderer
' that scales the whole page; the §9.8 usable-area work calibrated PAGINATION
' (which is correct) but nobody could see the RENDER scale until the output
' was measured with PyMuPDF.
'
' So the Location Map PDF is now written directly, byte by byte: one JPEG
' image per page placed at exactly the block geometry, the stamp and
' attribution as REAL vector text (crisper than printing, and searchable),
' and the site pushpin as vector art. No printer driver, no page breaks, no
' usable-area, no DPI - the same output on every machine, forever. The old
' print path survives in modMaps as the fallback when this module cannot run
' (WIA missing for PNG->JPEG, or an image's source file has gone away).
'
' Dependencies: ADODB.Stream (binary IO) + WIA (PNG->JPEG conversion), both
' inbox on every supported Windows - within the N1 zero-install rule.
'
' PDF specifics kept deliberately simple: PDF 1.4, uncompressed content
' streams, base-14 Helvetica / Helvetica-Bold with WinAnsi encoding (no font
' embedding), images as DCTDecode (JPEG pass-through) - the same trick that
' keeps jsPDF small in the web tool.

Private Const PAGE_W As Double = 792#     ' landscape Letter
Private Const PAGE_H As Double = 612#

' Block geometry mirrors the sheet design (§9.8): a MAP_PAGE_WIDTH_PTS x
' MAP_PAGE_HEIGHT_PTS block centered on the page -> 16pt side / 22pt
' top-bottom frame. All Y values below are PDF-style (origin bottom-left).
Private Const BLOCK_X As Double = 16#     ' (792-760)/2
Private Const BLOCK_Y As Double = 22#     ' (612-568)/2
Private Const IMG_INSET As Double = 1#

Private Const STAMP_FONT_PT As Double = 11#
Private Const STAMP_LINE_H As Double = 13.2

' JPEG encode quality for the aerials/screenshots embedded in the Location Map
' PDF. 80 is visually indistinguishable from 88 on aerial imagery (no fine text
' - the stamp/pin/attribution are drawn as vector, not baked into the JPEG) but
' meaningfully smaller. Every source is re-encoded at this quality, so a manual
' screenshot inserted as a large JPEG is compressed too (it used to pass through
' un-recompressed). Lower this toward ~70 for even smaller files.
Private Const JPEG_QUALITY As Long = 80

' Cap the embedded image width (px). Screenshots are often far larger than the
' ~760x568-pt page block needs; 1400 px across that block is ~133 DPI, ample for
' a location map, and downscaling to it is the biggest single lever on file size
' for oversized captures. Images already narrower than this are left as-is.
Private Const MAP_MAX_PX_W As Long = 1400

' One page's worth of gathered material.
Private Type PdfPageInfo
    StampText As String
    ImagePath As String      ' "" = no image (placeholder page)
    HasPin As Boolean
    AttrText As String       ' "" = no attribution line
End Type

' Touched by compile-check.ps1 to force-compile this module.
Public Function PdfSelfTest() As String
    PdfSelfTest = EscapePdfText("(ok)")
End Function

' ---- entry point ------------------------------------------------------------

' Build the COMBINED Location Map PDF (every page) straight from the Map Pages
' sheet's shapes + their source image files. True on success; False (with
' errMsg) on anything the fallback print path should handle instead.
Public Function BuildMapPdfDirect(ByVal wsMap As Worksheet, ByVal fullPath As String, _
        ByRef errMsg As String) As Boolean
    Dim nPages As Long, idx() As Long, i As Long
    nPages = MapPageCount(wsMap)
    If nPages < 1 Then errMsg = "no map pages": Exit Function
    ReDim idx(0 To nPages - 1)
    For i = 0 To nPages - 1
        idx(i) = i
    Next i
    BuildMapPdfDirect = BuildMapPdfForPages(wsMap, idx, fullPath, errMsg)
End Function

' Build a PDF for an explicit list of 0-based page indices (one page per index,
' in the given order). The combined export passes all pages; the per-site
' export passes a single index. True on success; False (with errMsg) otherwise.
Public Function BuildMapPdfForPages(ByVal wsMap As Worksheet, ByRef pageIdx() As Long, _
        ByVal fullPath As String, ByRef errMsg As String) As Boolean
    On Error GoTo Fail

    Dim cnt As Long
    cnt = UBound(pageIdx) - LBound(pageIdx) + 1
    If cnt < 1 Then errMsg = "no pages selected": Exit Function

    ' Gather per-page material first so any missing piece fails BEFORE a
    ' partial file is written.
    Dim pages() As PdfPageInfo, i As Long
    ReDim pages(0 To cnt - 1)
    For i = 0 To cnt - 1
        If Not GatherPage(wsMap, pageIdx(LBound(pageIdx) + i), pages(i), errMsg) Then Exit Function
    Next i

    ' Convert every image to JPEG bytes (pass-through when already JPEG).
    Dim jpegs() As Variant, dims() As Variant
    ReDim jpegs(0 To cnt - 1)
    ReDim dims(0 To cnt - 1)
    For i = 0 To cnt - 1
        If Len(pages(i).ImagePath) > 0 Then
            Dim jb() As Byte, jw As Long, jh As Long, jc As Long
            If Not JpegBytesFor(pages(i).ImagePath, jb, jw, jh, jc, errMsg) Then Exit Function
            jpegs(i) = jb
            dims(i) = Array(jw, jh, jc)
        End If
    Next i

    WriteMapPdf fullPath, pages, jpegs, dims
    BuildMapPdfForPages = True
    Exit Function
Fail:
    errMsg = Err.Description
    BuildMapPdfForPages = False
End Function

' Collect one page's stamp text, image file, pin + attribution presence.
Private Function GatherPage(ByVal wsMap As Worksheet, ByVal pageIdx As Long, _
        ByRef info As PdfPageInfo, ByRef errMsg As String) As Boolean
    Dim n As String
    n = CStr(pageIdx + 1)

    On Error Resume Next
    info.StampText = wsMap.Shapes("Textbox_Page_" & n).TextFrame.Characters.Text
    info.HasPin = Not wsMap.Shapes(MAP_PIN_PREFIX & n) Is Nothing
    info.AttrText = wsMap.Shapes(MAP_ATTR_PREFIX & n).TextFrame.Characters.Text
    On Error GoTo 0

    ' Image: only pages that HAVE a picture need a file; a placeholder page
    ' (no screenshot yet) renders as a white page with the stamp.
    Dim shp As Shape
    On Error Resume Next
    Set shp = wsMap.Shapes("MapImage_Page_" & n)
    On Error GoTo 0
    If shp Is Nothing Then
        info.ImagePath = ""
        GatherPage = True
        Exit Function
    End If

    info.ImagePath = ResolvePageImageFile(wsMap, shp, pageIdx)
    If Len(info.ImagePath) = 0 Then
        errMsg = "page " & n & ": image source file not found (re-run Fetch Imagery / Insert Images)"
        Exit Function
    End If
    GatherPage = True
End Function

' Where is page N's image on disk? 1) the path stashed on the shape at
' placement (AlternativeText); 2) the maps\ copies (Site_<no> / Page_<n>).
Private Function ResolvePageImageFile(ByVal wsMap As Worksheet, ByVal shp As Shape, _
        ByVal pageIdx As Long) As String
    Dim p As String
    p = shp.AlternativeText
    If Len(p) > 0 Then
        If Len(Dir$(p)) > 0 Then ResolvePageImageFile = p: Exit Function
    End If

    Dim folder As String, siteNo As String, siteRow As Long
    folder = ResolveOutputFolder() & "maps\"
    On Error Resume Next
    siteRow = CLng(Val(wsMap.Shapes("Textbox_Page_" & CStr(pageIdx + 1)).AlternativeText))
    If siteRow >= SITES_FIRST_DATA_ROW Then
        siteNo = Trim$(CStr(SitesSheet().Cells(siteRow, COL_SITENO).Value))
    End If
    On Error GoTo 0

    Dim exts As Variant, e As Variant
    exts = Array(".png", ".jpg", ".jpeg")
    If Len(siteNo) > 0 Then
        For Each e In exts
            p = folder & "Site_" & CleanFileName(siteNo) & e
            If Len(Dir$(p)) > 0 Then ResolvePageImageFile = p: Exit Function
        Next e
    End If
    For Each e In exts
        p = folder & "Page_" & CStr(pageIdx + 1) & e
        If Len(Dir$(p)) > 0 Then ResolvePageImageFile = p: Exit Function
    Next e
End Function

' ---- JPEG plumbing ----------------------------------------------------------

' Return JPEG bytes + pixel dimensions + component count for an image file,
' converting non-JPEGs via WIA. False (with errMsg) when conversion isn't
' possible - the caller falls back to the print path.
Private Function JpegBytesFor(ByVal imgPath As String, ByRef jpegBytes() As Byte, _
        ByRef pxW As Long, ByRef pxH As Long, ByRef comps As Long, _
        ByRef errMsg As String) As Boolean
    Dim ext As String, p As Long, jpgPath As String, madeTemp As Boolean, isJpeg As Boolean
    p = InStrRev(imgPath, ".")
    If p > 0 Then ext = LCase$(Mid$(imgPath, p))
    isJpeg = (ext = ".jpg" Or ext = ".jpeg")

    ' Re-encode EVERY source to a size-capped JPEG (downscale to MAP_MAX_PX_W +
    ' JPEG_QUALITY) so the PDF stays small even for big manual screenshots - a
    ' JPEG source used to be embedded whole. If WIA is unavailable an already-
    ' JPEG source still works (embedded as-is, just uncompressed); a PNG can't be
    ' embedded without WIA, so that still falls back to the print path.
    jpgPath = NormalizeToJpeg(imgPath)
    If Len(jpgPath) > 0 Then
        madeTemp = True
    ElseIf isJpeg Then
        jpgPath = imgPath
    Else
        errMsg = "PNG->JPEG conversion unavailable (WIA)"
        Exit Function
    End If

    jpegBytes = ReadFileBytes(jpgPath)
    If madeTemp Then
        On Error Resume Next
        Kill jpgPath
        On Error GoTo 0
    End If

    If Not ParseJpegHeader(jpegBytes, pxW, pxH, comps) Then
        errMsg = "could not parse JPEG dimensions for " & imgPath
        Exit Function
    End If
    If comps <> 1 And comps <> 3 Then
        errMsg = "unsupported JPEG color space (" & comps & " components)"
        Exit Function
    End If
    JpegBytesFor = True
End Function

' Any image WIA reads (PNG or JPEG) -> a size-capped JPEG temp file: downscale
' proportionally to MAP_MAX_PX_W when wider than the cap, then encode at
' JPEG_QUALITY. "" on failure (WIA unavailable / unreadable source).
Private Function NormalizeToJpeg(ByVal srcPath As String) As String
    On Error GoTo Fail
    Const WIA_FORMAT_JPEG As String = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
    Dim img As Object, ip As Object
    Set img = CreateObject("WIA.ImageFile")
    img.LoadFile srcPath
    Set ip = CreateObject("WIA.ImageProcess")

    ' Downscale first (proportionally) when wider than the cap - fewer pixels to
    ' encode is the biggest lever on size. The Scale filter fits within
    ' MaximumWidth x MaximumHeight preserving aspect; bounding height by the
    ' source height leaves WIDTH as the single binding constraint.
    If img.Width > MAP_MAX_PX_W Then
        ip.Filters.Add ip.FilterInfos("Scale").FilterID
        ip.Filters(ip.Filters.Count).Properties("MaximumWidth") = MAP_MAX_PX_W
        ip.Filters(ip.Filters.Count).Properties("MaximumHeight") = img.Height
    End If

    ip.Filters.Add ip.FilterInfos("Convert").FilterID
    ip.Filters(ip.Filters.Count).Properties("FormatID") = WIA_FORMAT_JPEG
    ip.Filters(ip.Filters.Count).Properties("Quality") = JPEG_QUALITY
    Set img = ip.Apply(img)

    Dim dst As String
    dst = Environ$("TEMP") & "\rr_pdf_" & Mid$(srcPath, InStrRev(srcPath, "\") + 1) & ".jpg"
    If Len(Dir$(dst)) > 0 Then Kill dst
    img.SaveFile dst
    NormalizeToJpeg = dst
    Exit Function
Fail:
    NormalizeToJpeg = ""
End Function

Private Function ReadFileBytes(ByVal path As String) As Byte()
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                     ' adTypeBinary
    stm.Open
    stm.LoadFromFile path
    ReadFileBytes = stm.Read
    stm.Close
End Function

' Scan the JPEG marker chain for the SOF frame header -> width/height/comps.
Private Function ParseJpegHeader(ByRef b() As Byte, ByRef pxW As Long, _
        ByRef pxH As Long, ByRef comps As Long) As Boolean
    Dim i As Long, marker As Long, segLen As Long, ub As Long
    ub = UBound(b)
    If ub < 3 Then Exit Function
    If b(0) <> &HFF Or b(1) <> &HD8 Then Exit Function
    i = 2
    Do While i + 3 <= ub
        If b(i) <> &HFF Then Exit Function
        marker = b(i + 1)
        Select Case marker
            Case &HC0, &HC1, &HC2, &HC3, &HC5, &HC6, &HC7, &HC9, &HCA, &HCB, &HCD, &HCE, &HCF
                If i + 9 > ub Then Exit Function
                pxH = CLng(b(i + 5)) * 256 + b(i + 6)
                pxW = CLng(b(i + 7)) * 256 + b(i + 8)
                comps = b(i + 9)
                ParseJpegHeader = (pxW > 0 And pxH > 0)
                Exit Function
            Case &HD8, &H1, &HD0 To &HD7    ' no-length markers
                i = i + 2
            Case Else
                segLen = CLng(b(i + 2)) * 256 + b(i + 3)
                i = i + 2 + segLen
        End Select
    Loop
End Function

' ---- PDF assembly -----------------------------------------------------------

' The whole file is assembled in one binary ADODB.Stream; object byte offsets
' are recorded as they are written (xref requirement).
Private Sub WriteMapPdf(ByVal fullPath As String, ByRef pages() As PdfPageInfo, _
        ByRef jpegs() As Variant, ByRef dims() As Variant)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open

    Dim nPages As Long
    nPages = UBound(pages) - LBound(pages) + 1

    ' Object numbering: 1 Catalog, 2 Pages, 3 F1, 4 F2, 5 ExtGState, then per
    ' page: content, page, (image). Compute image object numbers up front so
    ' page resources can reference them.
    Dim offsets() As Currency          ' Currency: exact integer beyond 2GB safe
    Dim totalObjs As Long, i As Long
    totalObjs = 5
    Dim contentNum() As Long, pageNum() As Long, imgNum() As Long
    ReDim contentNum(0 To nPages - 1)
    ReDim pageNum(0 To nPages - 1)
    ReDim imgNum(0 To nPages - 1)
    For i = 0 To nPages - 1
        totalObjs = totalObjs + 1: contentNum(i) = totalObjs
        totalObjs = totalObjs + 1: pageNum(i) = totalObjs
        If Len(pages(i).ImagePath) > 0 Then
            totalObjs = totalObjs + 1: imgNum(i) = totalObjs
        Else
            imgNum(i) = 0
        End If
    Next i
    ReDim offsets(1 To totalObjs)

    AppendStr stm, "%PDF-1.4" & vbLf
    ' Binary-content marker comment (spec recommendation).
    AppendBytes stm, MakeBytes(Array(&H25, &HE2, &HE3, &HCF, &HD3, &HA))

    ' 1: Catalog
    offsets(1) = stm.Position
    AppendStr stm, "1 0 obj" & vbLf & "<< /Type /Catalog /Pages 2 0 R >>" & vbLf & "endobj" & vbLf

    ' 2: Pages
    Dim kids As String
    For i = 0 To nPages - 1
        kids = kids & pageNum(i) & " 0 R "
    Next i
    offsets(2) = stm.Position
    AppendStr stm, "2 0 obj" & vbLf & "<< /Type /Pages /Kids [" & kids & "] /Count " & _
        nPages & " >>" & vbLf & "endobj" & vbLf

    ' 3 + 4: the base-14 fonts (no embedding needed).
    offsets(3) = stm.Position
    AppendStr stm, "3 0 obj" & vbLf & _
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>" & _
        vbLf & "endobj" & vbLf
    offsets(4) = stm.Position
    AppendStr stm, "4 0 obj" & vbLf & _
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>" & _
        vbLf & "endobj" & vbLf

    ' 5: ExtGState for the stamp/attribution boxes' slight translucency.
    offsets(5) = stm.Position
    AppendStr stm, "5 0 obj" & vbLf & "<< /Type /ExtGState /ca 0.85 /CA 1 >>" & vbLf & "endobj" & vbLf

    For i = 0 To nPages - 1
        ' Content stream.
        Dim content As String
        content = PageContentStream(pages(i), i, imgNum(i) > 0, dims(i))
        offsets(contentNum(i)) = stm.Position
        AppendStr stm, contentNum(i) & " 0 obj" & vbLf & "<< /Length " & Len(content) & _
            " >>" & vbLf & "stream" & vbLf & content & vbLf & "endstream" & vbLf & "endobj" & vbLf

        ' Page object.
        Dim res As String
        res = "/Font << /F1 3 0 R /F2 4 0 R >> /ExtGState << /GS1 5 0 R >>"
        If imgNum(i) > 0 Then res = res & " /XObject << /Im" & (i + 1) & " " & imgNum(i) & " 0 R >>"
        offsets(pageNum(i)) = stm.Position
        AppendStr stm, pageNum(i) & " 0 obj" & vbLf & "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " & _
            NumS(PAGE_W) & " " & NumS(PAGE_H) & "] /Resources << " & res & " >> /Contents " & _
            contentNum(i) & " 0 R >>" & vbLf & "endobj" & vbLf

        ' Image XObject (JPEG pass-through).
        If imgNum(i) > 0 Then
            Dim jb() As Byte, d As Variant, cs As String
            jb = jpegs(i)
            d = dims(i)
            cs = IIf(d(2) = 1, "/DeviceGray", "/DeviceRGB")
            offsets(imgNum(i)) = stm.Position
            AppendStr stm, imgNum(i) & " 0 obj" & vbLf & "<< /Type /XObject /Subtype /Image /Width " & _
                d(0) & " /Height " & d(1) & " /ColorSpace " & cs & _
                " /BitsPerComponent 8 /Filter /DCTDecode /Length " & (UBound(jb) - LBound(jb) + 1) & _
                " >>" & vbLf & "stream" & vbLf
            AppendBytes stm, jb
            AppendStr stm, vbLf & "endstream" & vbLf & "endobj" & vbLf
        End If
    Next i

    ' xref + trailer.
    Dim xrefPos As Currency
    xrefPos = stm.Position
    AppendStr stm, "xref" & vbLf & "0 " & (totalObjs + 1) & vbLf
    AppendStr stm, "0000000000 65535 f " & vbLf
    For i = 1 To totalObjs
        AppendStr stm, Format$(offsets(i), "0000000000") & " 00000 n " & vbLf
    Next i
    AppendStr stm, "trailer" & vbLf & "<< /Size " & (totalObjs + 1) & _
        " /Root 1 0 R >>" & vbLf & "startxref" & vbLf & CStr(CLng(xrefPos)) & vbLf & "%%EOF"

    stm.SaveToFile fullPath, 2       ' adSaveCreateOverWrite
    stm.Close
End Sub

' Drawing operations for one page: image (cover-cropped into the block via a
' clip path), stamp box + text, pushpin, attribution. The image resource is
' named /Im<pageIdx+1> - the SAME formula the page's resource dict uses.
Private Function PageContentStream(ByRef info As PdfPageInfo, ByVal pageIdx As Long, _
        ByVal hasImg As Boolean, ByVal d As Variant) As String
    Dim s As String
    Dim bw As Double, bh As Double
    bw = MAP_PAGE_WIDTH_PTS - IMG_INSET * 2      ' 758
    bh = MAP_PAGE_HEIGHT_PTS - IMG_INSET * 2     ' 566

    ' -- image, cover-scaled and clipped to the block (minus the 1pt inset) --
    If hasImg Then
        Dim iw As Double, ih As Double, scl As Double, dw As Double, dh As Double
        Dim ix As Double, iy As Double
        iw = d(0): ih = d(1)
        scl = bw / iw
        If bh / ih > scl Then scl = bh / ih      ' cover: larger scale wins
        dw = iw * scl: dh = ih * scl
        ix = BLOCK_X + IMG_INSET + (bw - dw) / 2#
        iy = BLOCK_Y + IMG_INSET + (bh - dh) / 2#
        s = s & "q" & vbLf
        s = s & NumS(BLOCK_X + IMG_INSET) & " " & NumS(BLOCK_Y + IMG_INSET) & " " & _
            NumS(bw) & " " & NumS(bh) & " re W n" & vbLf
        s = s & NumS(dw) & " 0 0 " & NumS(dh) & " " & NumS(ix) & " " & NumS(iy) & " cm" & vbLf
        s = s & "/Im" & (pageIdx + 1) & " Do" & vbLf
        s = s & "Q" & vbLf
    End If

    ' -- pushpin (fetched pages only): tip exactly at the page/site center --
    If info.HasPin Then s = s & PushpinOps()

    ' -- stamp: white box (slightly translucent) + text, top-left of block --
    If Len(Trim$(info.StampText)) > 0 Then s = s & StampOps(info.StampText)

    ' -- attribution line, bottom-left of block --
    If Len(Trim$(info.AttrText)) > 0 Then s = s & AttributionOps(info.AttrText)

    PageContentStream = s
End Function

Private Function PushpinOps() As String
    Dim s As String
    Dim cx As Double, cy As Double
    cx = PAGE_W / 2#
    cy = PAGE_H / 2#                 ' PDF origin bottom-left; page is symmetric

    ' Needle: downward grey triangle, tip on the site.
    s = s & "q" & vbLf
    s = s & "0.43 0.43 0.43 rg 0.27 0.27 0.27 RG 0.5 w" & vbLf
    s = s & NumS(cx - 2.25) & " " & NumS(cy + 13) & " m " & _
            NumS(cx + 2.25) & " " & NumS(cy + 13) & " l " & _
            NumS(cx) & " " & NumS(cy) & " l h B" & vbLf
    ' Head: yellow ball with darker outline.
    s = s & "1 0.8 0 rg 0.59 0.43 0 RG 0.75 w" & vbLf
    s = s & CircleOps(cx, cy + 19.5, 7.5) & " B" & vbLf
    ' Glint.
    s = s & "1 0.957 0.706 rg" & vbLf
    s = s & CircleOps(cx - 2.6, cy + 22.2, 2.1) & " f" & vbLf
    s = s & "Q" & vbLf
    PushpinOps = s
End Function

' Bezier circle path (not painted - caller appends f/B).
Private Function CircleOps(ByVal cx As Double, ByVal cy As Double, ByVal r As Double) As String
    Const K As Double = 0.552284749831
    Dim kr As Double
    kr = K * r
    CircleOps = _
        NumS(cx + r) & " " & NumS(cy) & " m " & _
        NumS(cx + r) & " " & NumS(cy + kr) & " " & NumS(cx + kr) & " " & NumS(cy + r) & " " & NumS(cx) & " " & NumS(cy + r) & " c " & _
        NumS(cx - kr) & " " & NumS(cy + r) & " " & NumS(cx - r) & " " & NumS(cy + kr) & " " & NumS(cx - r) & " " & NumS(cy) & " c " & _
        NumS(cx - r) & " " & NumS(cy - kr) & " " & NumS(cx - kr) & " " & NumS(cy - r) & " " & NumS(cx) & " " & NumS(cy - r) & " c " & _
        NumS(cx + kr) & " " & NumS(cy - r) & " " & NumS(cx + r) & " " & NumS(cy - kr) & " " & NumS(cx + r) & " " & NumS(cy) & " c h"
End Function

' Stamp: replicate the sheet textbox - MAP_TEXTBOX_WIDTH x MAP_TEXTBOX_HEIGHT
' at block-left+5 / block-top-5, white 85% box, grey hairline border, 11pt
' text with the first line bold.
Private Function StampOps(ByVal stampText As String) As String
    Dim bx As Double, byTop As Double, bxW As Double, bxH As Double
    bx = BLOCK_X + 5
    byTop = PAGE_H - BLOCK_Y - 5                 ' block top edge minus offset
    bxW = MAP_TEXTBOX_WIDTH
    bxH = MAP_TEXTBOX_HEIGHT

    Dim s As String
    s = s & "q /GS1 gs 1 1 1 rg 0.39 0.39 0.39 RG 0.5 w" & vbLf
    s = s & NumS(bx) & " " & NumS(byTop - bxH) & " " & NumS(bxW) & " " & NumS(bxH) & " re B" & vbLf
    s = s & "Q" & vbLf

    ' Text: margins 5pt left, 3pt top (mirrors the shape's TextFrame margins).
    Dim lines() As String, i As Long, y As Double, fnt As String
    lines = Split(Replace(stampText, vbCr, ""), vbLf)
    s = s & "q 0 0 0 rg BT" & vbLf
    y = byTop - 3 - STAMP_FONT_PT
    For i = LBound(lines) To UBound(lines)
        If y < byTop - bxH + 2 Then Exit For     ' never overflow the box
        fnt = IIf(i = LBound(lines), "/F2", "/F1")
        s = s & fnt & " " & NumS(STAMP_FONT_PT) & " Tf 1 0 0 1 " & _
            NumS(bx + 5) & " " & NumS(y) & " Tm (" & EscapePdfText(lines(i)) & ") Tj" & vbLf
        y = y - STAMP_LINE_H
    Next i
    s = s & "ET Q" & vbLf
    StampOps = s
End Function

' Attribution: small translucent white box + 7pt grey text, bottom-left of
' the block (mirrors modMapFetch.AddAttribution's shape).
Private Function AttributionOps(ByVal attrText As String) As String
    Dim s As String, w As Double
    w = 4 + Len(attrText) * 3.6                  ' rough 7pt Helvetica advance
    If w < 60 Then w = 60
    s = s & "q /GS1 gs 1 1 1 rg" & vbLf
    s = s & NumS(BLOCK_X + 2) & " " & NumS(BLOCK_Y + 2) & " " & NumS(w) & " 12 re f" & vbLf
    s = s & "0.24 0.24 0.24 rg BT /F1 7 Tf 1 0 0 1 " & NumS(BLOCK_X + 4) & " " & _
        NumS(BLOCK_Y + 5.5) & " Tm (" & EscapePdfText(attrText) & ") Tj ET" & vbLf
    s = s & "Q" & vbLf
    AttributionOps = s
End Function

' ---- low-level helpers ------------------------------------------------------

' Invariant number formatting for PDF operators ("." decimal, no grouping).
Private Function NumS(ByVal v As Double) As String
    NumS = Trim$(Str$(Int(v * 1000# + 0.5) / 1000#))
    If Left$(NumS, 1) = "." Then NumS = "0" & NumS
    If Left$(NumS, 2) = "-." Then NumS = "-0" & Mid$(NumS, 2)
End Function

' Escape a string for a PDF literal: backslash, parens; control chars dropped.
Public Function EscapePdfText(ByVal s As String) As String
    Dim i As Long, ch As String, out As String, code As Long
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch) And &HFFFF&
        Select Case ch
            Case "\": out = out & "\\"
            Case "(": out = out & "\("
            Case ")": out = out & "\)"
            Case Else
                If code >= 32 Then out = out & ch
        End Select
    Next i
    EscapePdfText = out
End Function

' Append a VBA string to the binary stream as ANSI (CP-1252 ~ WinAnsi) bytes.
Private Sub AppendStr(ByVal stm As Object, ByVal s As String)
    If Len(s) = 0 Then Exit Sub
    Dim b() As Byte
    b = StrConv(s, vbFromUnicode)
    stm.Write b
End Sub

Private Sub AppendBytes(ByVal stm As Object, ByRef b() As Byte)
    stm.Write b
End Sub

Private Function MakeBytes(ByVal arr As Variant) As Byte()
    Dim b() As Byte, i As Long
    ReDim b(0 To UBound(arr))
    For i = 0 To UBound(arr)
        b(i) = CByte(arr(i))
    Next i
    MakeBytes = b
End Function
