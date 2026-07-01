Attribute VB_Name = "modBuild"
Option Explicit

' RoadReviewer V1 - workbook builder (verification step 5.2).
' Run BuildWorkbook once on a fresh workbook, then save as RoadReviewer.xlsm.
' Re-running is safe: layout/control sheets are rebuilt and the Sites table's
' derived columns (hyperlinks/validation/formatting) are re-applied, but data
' the inspector typed into the Sites table is preserved.

Private Const CLR_HEADER As Long = 5197615      ' dark slate (RGB 47,79,79-ish)
Private Const CLR_BTN As Long = 12419407         ' steel blue
Private Const CLR_BTN_GO As Long = 4563272        ' green
' VBA Long color = R + G*256 + B*65536. Calibrated to Excel's standard
' light-red / light-green / light-yellow conditional-format swatches.
Private Const CLR_FEDAID As Long = 13551615      ' RGB(255,199,206) — "Federal aid" rows
Private Const CLR_NONFEDAID As Long = 13561798   ' RGB(198,239,206) — "Non-federal aid" rows
Private Const CLR_REVIEW As Long = 10284031      ' RGB(255,235,156) — "Review" rows

Public Sub BuildWorkbook()
    Dim hadSites As Boolean
    hadSites = SheetExists(SH_SITES)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Fail

    BuildSites                ' first - preserves data if present
    BuildSetup
    BuildClassifySheet
    BuildImagerySheet
    BuildMapsSheet
    BuildHome                 ' built last, then activated

    RemoveStrayDefaultSheets
    OrderSheets

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    SheetByName(SH_HOME).Activate

    ' Skip the confirmation MsgBox when driven headlessly by the build script
    ' (which sets gHeadless before calling). The in-Excel "Build / Reset
    ' Workbook" button leaves gHeadless=False and still shows the dialog.
    If Not gHeadless Then
        MsgBox "RoadReviewer workbook built." & vbCrLf & vbCrLf & _
            IIf(hadSites, "Existing Sites data was preserved.", "Start by filling in the Setup sheet, then add points on the Sites sheet.") & vbCrLf & _
            "Remember to save as a macro-enabled workbook (.xlsm).", _
            vbInformation, "RoadReviewer"
    End If
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If gHeadless Then
        ' Re-raise so the automation host sees the failure.
        Err.Raise Err.Number, "BuildWorkbook", Err.Description
    Else
        MsgBox "Build failed: " & Err.Description, vbCritical, "RoadReviewer"
    End If
End Sub

' ---- sheet plumbing -------------------------------------------------------

' Delete-and-recreate a control sheet (no user data lives on these).
Private Function FreshSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    Set ws = SheetByName(sheetName)
    If Not ws Is Nothing Then ws.Delete
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    Set FreshSheet = ws
End Function

' Delete leftover blank default sheets (Sheet1/Sheet2/...) that aren't ours.
Private Sub RemoveStrayDefaultSheets()
    Dim ws As Worksheet, i As Long
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        Select Case ws.Name
            Case SH_HOME, SH_SETUP, SH_SITES, SH_CLASSIFY, SH_IMAGERY, SH_MAPS
                ' keep
            Case Else
                If ThisWorkbook.Worksheets.Count > 1 Then
                    If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then ws.Delete
                End If
        End Select
    Next i
End Sub

Private Sub OrderSheets()
    MoveSheet SH_HOME, 1
    MoveSheet SH_SETUP, 2
    MoveSheet SH_SITES, 3
    MoveSheet SH_CLASSIFY, 4
    MoveSheet SH_IMAGERY, 5
    MoveSheet SH_MAPS, 6
End Sub

Private Sub MoveSheet(ByVal sheetName As String, ByVal pos As Long)
    Dim ws As Worksheet
    Set ws = SheetByName(sheetName)
    If ws Is Nothing Then Exit Sub
    If pos >= ThisWorkbook.Worksheets.Count Then
        ws.Move After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Else
        ws.Move Before:=ThisWorkbook.Worksheets(pos)
    End If
End Sub

' Add a flat rounded-rectangle button wired to a macro via OnAction.
Private Function AddButton(ByVal ws As Worksheet, ByVal leftPt As Single, ByVal topPt As Single, _
        ByVal widthPt As Single, ByVal heightPt As Single, ByVal caption As String, _
        ByVal macroName As String, Optional ByVal fillColor As Long = CLR_BTN) As Shape
    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPt, topPt, widthPt, heightPt)
    sh.Fill.ForeColor.RGB = fillColor
    sh.Line.Visible = msoFalse
    sh.Shadow.Visible = msoFalse
    With sh.TextFrame2.TextRange
        .Text = caption
        .Font.Size = 11
        .Font.Bold = msoTrue
        .Font.Fill.ForeColor.RGB = vbWhite
        .ParagraphFormat.Alignment = msoAlignCenter
    End With
    sh.TextFrame2.VerticalAnchor = msoAnchorMiddle
    sh.TextFrame2.HorizontalAnchor = msoAnchorCenter
    sh.OnAction = macroName
    Set AddButton = sh
End Function

' DisplayGridlines lives on Window, not Worksheet. To hide gridlines on a
' specific sheet we activate it first, then flip the Window property.
' On Error Resume Next handles the rare COM case where ActiveWindow is Nothing
' (Excel driven by an automation host with no visible workbook window).
Private Sub HideGridlines(ByVal ws As Worksheet)
    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0
End Sub

Private Sub TitleBlock(ByVal ws As Worksheet, ByVal title As String, ByVal subtitle As String)
    With ws.Range("B2")
        .Value = title
        .Font.Size = 20
        .Font.Bold = True
    End With
    With ws.Range("B3")
        .Value = subtitle
        .Font.Size = 11
        .Font.Italic = True
        .Font.Color = RGB(90, 90, 90)
    End With
End Sub

' ---- Home -----------------------------------------------------------------

Private Sub BuildHome()
    Dim ws As Worksheet
    Set ws = FreshSheet(SH_HOME)
    ws.Cells.Interior.Color = RGB(245, 247, 249)
    ws.Columns("A").ColumnWidth = 2
    TitleBlock ws, "RoadReviewer", "FEMA Public Assistance Site Inspector toolkit"

    ws.Range("B5").Value = "Workflow - run top to bottom:"
    ws.Range("B5").Font.Bold = True

    AddButton ws, 18, 110, 230, 34, "Set Up Job (WO / DI / State / Folder)", "GoSetup", CLR_BTN
    AddButton ws, 18, 150, 230, 34, "Enter Sites (addresses or lat/lon)", "GoSites", CLR_BTN
    AddButton ws, 320, 150, 210, 34, "Geocode Addresses -> lat/lon", "GeocodeAddresses", RGB(110, 110, 110)
    AddButton ws, 18, 200, 230, 40, "1.  Classify Roads", "GoClassify", CLR_BTN_GO
    AddButton ws, 18, 246, 230, 40, "2.  Review Imagery", "GoImagery", CLR_BTN_GO
    AddButton ws, 18, 292, 230, 40, "3.  Maps & FIRMettes", "GoMaps", CLR_BTN_GO

    ws.Range("E11").Value = "Each workflow reads from the shared Sites table. " & _
        "Enter every point once on the Sites sheet; the workflows fill in the rest."
    ws.Range("E11").Font.Color = RGB(70, 70, 70)
    ws.Range("E13").Value = "Need to rebuild the layout? (Sites data is preserved.)"
    AddButton ws, 320, 200, 180, 28, "Build / Reset Workbook", "BuildWorkbook", RGB(150, 150, 150)

    ws.Range("E16").Value = "Road classification is wired for Michigan, Indiana and Wisconsin in V1. " & _
        "ACUB (urban-boundary) lookup is nationwide."
    ws.Range("E16").Font.Italic = True
    ws.Range("E16").Font.Color = RGB(70, 70, 70)

    HideGridlines ws
    ws.Tab.Color = RGB(47, 79, 79)
End Sub

Public Sub GoSetup(): SheetByName(SH_SETUP).Activate: End Sub
Public Sub GoSites(): SheetByName(SH_SITES).Activate: End Sub
Public Sub GoClassify(): SheetByName(SH_CLASSIFY).Activate: End Sub
Public Sub GoImagery(): SheetByName(SH_IMAGERY).Activate: End Sub
Public Sub GoMaps(): SheetByName(SH_MAPS).Activate: End Sub

' ---- Setup ----------------------------------------------------------------

Private Sub BuildSetup()
    Dim ws As Worksheet
    Set ws = FreshSheet(SH_SETUP)
    HideGridlines ws
    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 60
    TitleBlock ws, "Setup", "Job-wide values. WO/DI default onto every Sites row; override per row if needed."

    LabelValue ws, 5, "Work Order (WO #)", NR_WO, ""
    LabelValue ws, 6, "Disaster Incident (DI #)", NR_DI, ""
    LabelValue ws, 7, "Disaster Number", NR_DISASTER, ""
    LabelValue ws, 8, "Applicant", NR_APPLICANT, ""
    LabelValue ws, 9, "State", NR_STATE, "MI"
    LabelValue ws, 10, "Output Folder", NR_OUTFOLDER, ""
    LabelValue ws, 11, "AGOL Webmap URL (optional)", NR_AGOLMAP, ""
    LabelValue ws, 12, "Search buffer (feet)", NR_BUFFER, CStr(DEFAULT_BUFFER_FEET)

    ' State dropdown (F8).
    With ws.Range("B9").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=STATE_LIST
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' Buffer must be a whole number between 1 and 1000 ft. Out-of-range
    ' values fall back to DEFAULT_BUFFER_FEET via modClassify.BufferFeet,
    ' but the validation catches mistakes at entry time.
    With ws.Range("B12").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="1", Formula2:="1000"
        .IgnoreBlank = True
        .ErrorMessage = "Enter a whole number between 1 and 1000 ft."
        .ShowError = True
    End With

    AddButton ws, ws.Range("D10").Left, ws.Range("B10").Top - 2, 150, 22, "Browse for folder...", "SelectOutputFolder"

    ws.Range("A14").Value = "Michigan, Indiana and Wisconsin road-class layers are wired in V1. Other states still run the ACUB check."
    ws.Range("A14").Font.Italic = True
    ws.Range("A14").Font.Color = RGB(90, 90, 90)
    ws.Range("A15").Value = "AGOL Webmap URL is optional. Paste a https://www.arcgis.com/apps/mapviewer/...?webmap=<id> URL " & _
        "(or your org's equivalent) to enable the AGOL Map column and the Send-to-AGOL button."
    ws.Range("A15").Font.Italic = True
    ws.Range("A15").Font.Color = RGB(90, 90, 90)
    ws.Range("A16").Value = "Search buffer is the radius used when no road segment intersects the exact point. 200 ft is a good default; lower it for dense urban grids, raise it for sparse rural networks."
    ws.Range("A16").Font.Italic = True
    ws.Range("A16").Font.Color = RGB(90, 90, 90)
    ws.Tab.Color = RGB(70, 130, 180)
End Sub

Private Sub LabelValue(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
        ByVal namedRange As String, ByVal defaultVal As String)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 1).Font.Bold = True
    With ws.Cells(r, 2)
        .Interior.Color = RGB(255, 255, 204)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        If Len(defaultVal) > 0 Then .Value = defaultVal
    End With
    AddNameForCell ws.Cells(r, 2), namedRange
End Sub

Private Sub AddNameForCell(ByVal cell As Range, ByVal nm As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & cell.Worksheet.Name & "'!" & cell.Address(True, True)
End Sub

' ---- Sites (the shared table) --------------------------------------------

Private Sub BuildSites()
    Dim ws As Worksheet, isNew As Boolean
    Set ws = SheetByName(SH_SITES)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_SITES
        isNew = True
    End If

    WriteSitesHeader ws
    FillSitesFormulas ws
    ApplySitesValidation ws
    ApplySitesFormatting ws

    ws.Rows(SITES_HEADER_ROW).Font.Bold = True
    ws.Activate
    ws.Range("A2").Select
    On Error Resume Next
    ' ActiveWindow can be Nothing when Excel is driven headless by COM.
    ' Skipping the freeze in that path is fine - the in-Excel "Build / Reset"
    ' button still pins the header row.
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    ws.Tab.Color = RGB(60, 60, 60)
End Sub

Private Sub WriteSitesHeader(ByVal ws As Worksheet)
    Dim h() As String
    ReDim h(1 To COL_LAST)
    h(COL_WO) = "WO #"
    h(COL_DI) = "DI #"
    h(COL_SITENO) = "Site #"
    h(COL_SITENAME) = "Site Name"
    h(COL_ADDRESS) = "Address"
    h(COL_LAT) = "Latitude"
    h(COL_LON) = "Longitude"
    h(COL_CATEGORY) = "Category"
    h(COL_DESC) = "Description"
    h(COL_COSTS) = "Costs"
    h(COL_WORKCOMP) = "Work Completion"
    h(COL_GEOCODE) = "Geocode Status"
    h(COL_GMAP) = "Google Maps"
    h(COL_STREETVIEW) = "Street View"
    h(COL_BING) = "Bing"
    h(COL_FEMAVIEW) = "FEMA Viewer"
    h(COL_FIRMPORTAL) = "FIRMette Portal"
    h(COL_NFCMAP) = "MDOT NFC Map"
    h(COL_CLASS) = "FHWA Class"
    h(COL_URBANRURAL) = "Urban/Rural"
    h(COL_ACUBNAME) = "ACUB Name"
    h(COL_ROADNAME) = "Road Name"
    h(COL_STREET) = "Street Name"
    h(COL_ELIGIBILITY) = "Federal Aid Status"
    h(COL_FIRMSTATUS) = "FIRMette Status"
    h(COL_MAPSTATUS) = "Map Status"
    h(COL_AGOLMAP) = "AGOL Map"

    Dim c As Long
    For c = 1 To COL_LAST
        ws.Cells(SITES_HEADER_ROW, c).Value = h(c)
    Next c
End Sub

' Re-apply every link formula on the Sites sheet. Called by BuildWorkbook as
' part of the initial build, and exposed as a Public sub so automation
' hosts and the in-Excel "Build / Reset Workbook" path can restore the
' link columns after a clear without redoing the whole sheet build.
Public Sub RefreshSitesFormulas()
    FillSitesFormulas SitesSheet()
End Sub

' Hyperlink columns are pure formulas off lat/lon (zero macro cost). Assigning
' the formula to a multi-row range makes Excel adjust the relative refs per row.
Private Sub FillSitesFormulas(ByVal ws As Worksheet)
    Dim r1 As Long, r2 As Long, latC As String, lonC As String
    r1 = SITES_FIRST_DATA_ROW
    r2 = SITES_FIRST_DATA_ROW + SITES_FORMULA_ROWS - 1
    latC = "$" & ColLetter(COL_LAT) & r1
    lonC = "$" & ColLetter(COL_LON) & r1

    SetLinkFormula ws, COL_GMAP, r1, r2, URL_GMAP, "Map", latC, lonC
    SetLinkFormula ws, COL_STREETVIEW, r1, r2, URL_STREETVIEW, "Street", latC, lonC
    SetLinkFormula ws, COL_BING, r1, r2, URL_BING, "Bing", latC, lonC
    SetLinkFormula ws, COL_FEMAVIEW, r1, r2, URL_FEMAVIEW, "FEMA", latC, lonC
    SetLinkFormula ws, COL_FIRMPORTAL, r1, r2, URL_FIRMPORTAL, "FIRMette", latC, lonC
    ' Per-row "Open in map" (F11). Now points at the ArcGIS Online Map
    ' Viewer with the MDOT NFC FeatureServer side-loaded via the `url=`
    ' query param — bypasses the Experience app popup the inspector
    ' previously had to dismiss before they could see their point.
    SetLinkFormula ws, COL_NFCMAP, r1, r2, URL_NFC_MAPVIEW, "Open", latC, lonC, True
    ' The AGOL Map column is driven by the inspector's own webmap URL
    ' on Setup. The formula handles all three "blank" states (no URL
    ' set, missing coords) and stitches center/level/marker query
    ' params onto whatever URL was pasted.
    SetAgolMapFormula ws, r1, r2, latC, lonC
End Sub

' The AGOL formula is bespoke enough (depends on the dynamic Setup URL,
' picks ? vs & based on whether the URL already has query params) that
' it doesn't fit SetLinkFormula. Built as its own helper.
Private Sub SetAgolMapFormula(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
        ByVal latC As String, ByVal lonC As String)
    Dim sep As String, urlExpr As String, f As String
    ' If the user's pasted URL already contains "?", join with "&", else with "?".
    sep = "IF(ISNUMBER(FIND(""?""," & NR_AGOLMAP & ")),""&"",""?"")"
    urlExpr = NR_AGOLMAP & "&" & sep & _
        "&""center=""&" & lonC & "&"",""&" & latC & _
        "&""&level=16&marker=""&" & lonC & "&"",""&" & latC
    ' Blank cell when either the AGOL URL isn't set OR the row has no coords.
    f = "=IF(OR(" & NR_AGOLMAP & "="""", " & latC & "="""", " & lonC & "=""""),""""," & _
        "HYPERLINK(" & urlExpr & ",""Open""))"
    ws.Range(ws.Cells(r1, COL_AGOLMAP), ws.Cells(r2, COL_AGOLMAP)).Formula = f
End Sub

Private Sub SetLinkFormula(ByVal ws As Worksheet, ByVal col As Long, ByVal r1 As Long, ByVal r2 As Long, _
        ByVal urlTemplate As String, ByVal friendly As String, ByVal latC As String, ByVal lonC As String, _
        Optional ByVal needsCoords As Boolean = True)
    Dim urlExpr As String, f As String
    ' Build a formula expression that substitutes lat/lon into the template.
    urlExpr = """" & Replace(Replace(urlTemplate, "{LAT}", """&" & latC & "&"""), "{LON}", """&" & lonC & "&""") & """"
    If needsCoords Then
        f = "=IF(OR(" & latC & "="""" ," & lonC & "=""""),"""",HYPERLINK(" & urlExpr & ",""" & friendly & """))"
    Else
        f = "=HYPERLINK(" & urlExpr & ",""" & friendly & """)"
    End If
    ws.Range(ws.Cells(r1, col), ws.Cells(r2, col)).Formula = f
End Sub

Private Sub ApplySitesValidation(ByVal ws As Worksheet)
    Dim r2 As Long
    r2 = SITES_FIRST_DATA_ROW + SITES_FORMULA_ROWS - 1
    AddDecimalValidation ws.Range(ws.Cells(SITES_FIRST_DATA_ROW, COL_LAT), ws.Cells(r2, COL_LAT)), -90, 90, "Latitude must be between -90 and 90."
    AddDecimalValidation ws.Range(ws.Cells(SITES_FIRST_DATA_ROW, COL_LON), ws.Cells(r2, COL_LON)), -180, 180, "Longitude must be between -180 and 180."
End Sub

Private Sub AddDecimalValidation(ByVal rng As Range, ByVal lo As Double, ByVal hi As Double, ByVal msg As String)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:=CStr(lo), Formula2:=CStr(hi)
        .IgnoreBlank = True
        .ErrorMessage = msg
        .ShowError = True
    End With
End Sub

Private Sub ApplySitesFormatting(ByVal ws As Worksheet)
    Dim r2 As Long, eligCol As String
    r2 = SITES_FIRST_DATA_ROW + SITES_FORMULA_ROWS - 1

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11
    With ws.Rows(SITES_HEADER_ROW)
        .Interior.Color = CLR_HEADER
        .Font.Color = vbWhite
        .Font.Bold = True
    End With

    ' Sensible widths.
    ws.Columns(COL_SITENAME).ColumnWidth = 22
    ws.Columns(COL_ADDRESS).ColumnWidth = 28
    ws.Columns(COL_DESC).ColumnWidth = 26
    ws.Columns(COL_COSTS).ColumnWidth = 14
    ws.Columns(COL_WORKCOMP).ColumnWidth = 16
    ws.Columns(COL_CLASS).ColumnWidth = 16
    ws.Columns(COL_ACUBNAME).ColumnWidth = 16
    ws.Columns(COL_ROADNAME).ColumnWidth = 16
    ws.Columns(COL_STREET).ColumnWidth = 22
    ws.Columns(COL_ELIGIBILITY).ColumnWidth = 24
    ws.Columns(COL_GEOCODE).ColumnWidth = 16
    ws.Columns(COL_FIRMSTATUS).ColumnWidth = 18
    ws.Columns(COL_MAPSTATUS).ColumnWidth = 16
    ws.Columns(COL_AGOLMAP).ColumnWidth = 10

    ' Tri-state highlight on the Federal Aid Status column:
    '   red    — cell starts with "Federal aid"  (federal-aid road)
    '   green  — cell starts with "Non-federal aid"
    '   yellow — cell starts with "Review" (non-certified class or no
    '            road found within the Setup search-buffer radius)
    ' "Non-federal aid" intentionally tests for the literal prefix
    ' (LEFT … 15) because a substring search for "Federal aid" would
    ' also match "Non-federal aid". Order matters when format rules
    ' overlap: in Excel the FIRST matching rule wins, so non-federal
    ' is checked before federal.
    eligCol = "$" & ColLetter(COL_ELIGIBILITY) & SITES_FIRST_DATA_ROW
    With ws.Range(ws.Cells(SITES_FIRST_DATA_ROW, COL_CLASS), ws.Cells(r2, COL_ELIGIBILITY))
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlExpression, _
            Formula1:="=LEFT(" & eligCol & ",15)=""Non-federal aid"""
        .FormatConditions(.FormatConditions.Count).Interior.Color = CLR_NONFEDAID
        .FormatConditions.Add Type:=xlExpression, _
            Formula1:="=LEFT(" & eligCol & ",11)=""Federal aid"""
        .FormatConditions(.FormatConditions.Count).Interior.Color = CLR_FEDAID
        .FormatConditions.Add Type:=xlExpression, _
            Formula1:="=LEFT(" & eligCol & ",6)=""Review"""
        .FormatConditions(.FormatConditions.Count).Interior.Color = CLR_REVIEW
    End With

    On Error Resume Next
    ws.Range(ws.Cells(SITES_HEADER_ROW, 1), ws.Cells(SITES_HEADER_ROW, COL_LAST)).AutoFilter
    On Error GoTo 0
End Sub

' ---- Workflow control sheets ---------------------------------------------

Private Sub BuildClassifySheet()
    Dim ws As Worksheet
    Set ws = FreshSheet(SH_CLASSIFY)
    HideGridlines ws
    TitleBlock ws, "1.  Classify Roads", "Look up FHWA functional class + ACUB urban boundary for every Sites row."
    Bullets ws, 5, Array( _
        "Reads Latitude/Longitude from the Sites table.", _
        "Writes FHWA Class, Urban/Rural, ACUB Name, Road Name and Eligibility back to each row.", _
        "Federal aid rows (Urban Minor Collector or greater) are highlighted red, non-federal aid green, review yellow.", _
        "Michigan, Indiana and Wisconsin road class are wired in V1; other states still get the ACUB check.")
    AddButton ws, 18, 150, 200, 34, "Classify All Rows", "ClassifyAllRows", CLR_BTN_GO
    AddButton ws, 230, 150, 200, 34, "Re-run Failed Rows", "ReRunFailedClassifications", CLR_BTN
    ws.Tab.Color = RGB(46, 139, 87)
End Sub

Private Sub BuildImagerySheet()
    Dim ws As Worksheet
    Set ws = FreshSheet(SH_IMAGERY)
    HideGridlines ws
    TitleBlock ws, "2.  Review Imagery", "Open a curated set of imagery sources for the selected Sites row(s)."
    Bullets ws, 5, Array( _
        "Select one or more rows in the Sites table first.", _
        "Opens Google Maps, Street View, Bing aerial, FEMA Map Viewer for each point.", _
        "Each opens in your default browser; review pre-disaster condition there.")
    AddButton ws, 18, 140, 280, 36, "Open Imagery for Selected Row(s)", "OpenImageryForSelection", CLR_BTN_GO
    ws.Tab.Color = RGB(70, 130, 180)
End Sub

Private Sub BuildMapsSheet()
    Dim ws As Worksheet
    Set ws = FreshSheet(SH_MAPS)
    HideGridlines ws
    TitleBlock ws, "3.  Maps & FIRMettes", "Batch FEMA FIRMette download, per-site map pages, and exports."
    Bullets ws, 5, Array( _
        "Set WO/DI/Disaster and the output folder on the Setup sheet first.", _
        "FIRMettes download as PDFs named per site; status lands in the Sites table.", _
        "Map pages give one printable page per site for a pasted screenshot.")
    AddButton ws, 18, 140, 190, 32, "Download FIRMettes", "DownloadFirmettes", CLR_BTN_GO
    AddButton ws, 218, 140, 190, 32, "Re-run Failed FIRMettes", "ReRunFailedFirmettes", CLR_BTN
    AddButton ws, 18, 180, 190, 32, "Prepare Map Pages", "PrepareMapPages", CLR_BTN
    AddButton ws, 218, 180, 190, 32, "Export Combined Map PDF", "ExportCombinedMapPdf", CLR_BTN
    AddButton ws, 18, 220, 190, 32, "Export Sites to KML", "ExportSitesToKML", CLR_BTN
    AddButton ws, 218, 220, 190, 32, "Export Sites Table (CSV)", "ExportSitesCsv", CLR_BTN
    AddButton ws, 18, 260, 390, 32, "Send Sites to AGOL Map (KML + open webmap)", "SendSitesToAgolMap", CLR_BTN
    ws.Tab.Color = RGB(184, 134, 11)
End Sub

Private Sub Bullets(ByVal ws As Worksheet, ByVal startRow As Long, ByVal items As Variant)
    Dim i As Long
    For i = LBound(items) To UBound(items)
        ws.Cells(startRow + i, 2).Value = ChrW(8226) & "  " & items(i)
        ws.Cells(startRow + i, 2).Font.Color = RGB(70, 70, 70)
    Next i
End Sub
