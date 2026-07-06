Attribute VB_Name = "modBuild"
Option Explicit

' RoadReviewer / Site Inspector Review Tool - workbook builder.
' Run BuildWorkbook once on a fresh workbook (build.ps1 does this after
' calling SetProduct), then save as .xlsm. Re-running is safe: layout/control
' sheets are rebuilt and the Sites table's derived columns (hyperlinks/
' validation/formatting) are re-applied, but data typed into the Sites table
' is preserved.
'
' Both products share the same three-sheet shape:
'   Start Here - job inputs + every action button (no nav-only buttons)
'   Sites      - the shared table (toolbar row + header + 500 formula rows)
'   Sources    - per-state data-source citations + quirks (modSources)
' The standard product simply builds fewer inputs/buttons and hides the
' inspector-only Sites columns.

Private Const CLR_HEADER As Long = 5197615      ' dark slate (RGB 47,79,79-ish)
Private Const CLR_BTN As Long = 12419407         ' steel blue
Private Const CLR_BTN_GO As Long = 4563272        ' green
' VBA Long color = R + G*256 + B*65536. Calibrated to Excel's standard
' light-red / light-green / light-yellow conditional-format swatches.
Private Const CLR_FEDAID As Long = 13551615      ' RGB(255,199,206) — "Federal aid" rows
Private Const CLR_NONFEDAID As Long = 13561798   ' RGB(198,239,206) — "Non-federal aid" rows
Private Const CLR_REVIEW As Long = 10284031      ' RGB(255,235,156) — "Review" rows
' Sites column guidance tints: yellow = the user types here, grey = a
' workflow writes here. Link columns stay white (they're formulas).
Private Const CLR_INPUT As Long = 13434879       ' RGB(255,255,204) light yellow
Private Const CLR_RESULT As Long = 15921906      ' RGB(242,242,242) light grey

Public Sub BuildWorkbook()
    Dim hadSites As Boolean
    hadSites = SheetExists(SH_SITES)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo Fail

    BuildSites                ' first - preserves data if present
    BuildSourcesSheet         ' modSources
    BuildStartHere            ' built last, then activated

    RemoveStrayDefaultSheets
    OrderSheets

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    SheetByName(SH_START).Activate

    ' Skip the confirmation MsgBox when driven headlessly by the build script
    ' (which sets gHeadless before calling). The in-Excel "Build / Reset
    ' Workbook" button leaves gHeadless=False and still shows the dialog.
    If Not gHeadless Then
        MsgBox ProductTitle() & " workbook built." & vbCrLf & vbCrLf & _
            IIf(hadSites, "Existing Sites data was preserved.", _
                "Pick your state on the Start Here sheet, then add points on the Sites sheet.") & vbCrLf & _
            "Remember to save as a macro-enabled workbook (.xlsm).", _
            vbInformation, ProductTitle()
    End If
    Exit Sub
Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If gHeadless Then
        ' Re-raise so the automation host sees the failure.
        Err.Raise Err.Number, "BuildWorkbook", Err.Description
    Else
        MsgBox "Build failed: " & Err.Description, vbCritical, ProductTitle()
    End If
End Sub

' ---- sheet plumbing -------------------------------------------------------

' Delete-and-recreate a control sheet (no user data lives on these).
' Public because modSources uses it for the Sources sheet.
Public Function FreshSheet(ByVal sheetName As String) As Worksheet
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
            Case SH_START, SH_SITES, SH_SOURCES, SH_MAPPAGES
                ' keep
            Case Else
                If ThisWorkbook.Worksheets.Count > 1 Then
                    If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then ws.Delete
                End If
        End Select
    Next i
End Sub

Private Sub OrderSheets()
    MoveSheet SH_START, 1
    MoveSheet SH_SITES, 2
    MoveSheet SH_SOURCES, 3
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
' Public because modSources uses it for the Sources sheet.
Public Sub HideGridlines(ByVal ws As Worksheet)
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

Private Sub NoteLine(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    ws.Cells(r, 2).Value = txt
    ws.Cells(r, 2).Font.Italic = True
    ws.Cells(r, 2).Font.Color = RGB(90, 90, 90)
End Sub

Private Sub StepLine(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    ws.Cells(r, 2).Value = txt
    ws.Cells(r, 2).Font.Color = RGB(70, 70, 70)
End Sub

Private Sub SectionLabel(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    ws.Cells(r, 2).Value = txt
    ws.Cells(r, 2).Font.Bold = True
    ws.Cells(r, 2).Font.Size = 12
End Sub

' Prominent red-bordered disclaimer box spanning B:C over `rowCount` rows.
' Same wording on both products - this is the "not authoritative" contract
' the user asked to have front and center. Kept in sync with modSources'
' echo of it and with web/index.html's on-page disclaimer.
Private Sub DisclaimerBlock(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal rowCount As Long)
    Dim rng As Range, r As Long, hdr As String, body As String
    hdr = "IMPORTANT - NOT AN AUTHORITATIVE FHWA OR ELIGIBILITY DETERMINATION"
    body = "This tool does NOT authoritatively identify FHWA federal-aid roads. It flags high-probability " & _
        "candidates for a person to review, and may miss or mis-tag roads. It is not an authoritative source " & _
        "for FHWA functional classification. EVERY coordinate must be verified by a human against the official " & _
        "source map - use each row's NFC Map link and the Sources tab. Results are informational only and do " & _
        "NOT constitute a federal-aid, funding, or eligibility determination. A point on or near an urban/rural " & _
        "boundary is deliberately treated as Urban (within the search buffer; see the Sources tab) so boundary " & _
        "roads are not missed - always confirm these manually."
    Set rng = ws.Range(ws.Cells(firstRow, 2), ws.Cells(firstRow + rowCount - 1, 3))
    rng.Merge
    rng.Value = hdr & vbLf & body
    rng.WrapText = True
    rng.HorizontalAlignment = xlLeft
    rng.VerticalAlignment = xlTop
    rng.Font.Color = RGB(150, 0, 0)
    rng.Interior.Color = RGB(255, 238, 238)
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = RGB(192, 0, 0)
        .Weight = xlMedium
    End With
    ' Bold just the header line (merged text lives in the top-left cell).
    ws.Cells(firstRow, 2).Characters(1, Len(hdr)).Font.Bold = True
    For r = firstRow To firstRow + rowCount - 1
        ws.Rows(r).RowHeight = 21
    Next r
End Sub

' Small grey build/version stamp so a shared copy is traceable to its PR.
Private Sub VersionLabel(ByVal ws As Worksheet, ByVal r As Long)
    With ws.Cells(r, 2)
        .Value = ProductTitle() & "  -  " & BUILD_REFERENCE
        .Font.Size = 9
        .Font.Color = RGB(130, 130, 130)
    End With
End Sub

' ---- Start Here -----------------------------------------------------------

Private Sub BuildStartHere()
    Dim ws As Worksheet
    Set ws = FreshSheet(SH_START)
    ws.Cells.Interior.Color = RGB(245, 247, 249)
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 26
    ws.Columns("C").ColumnWidth = 62
    HideGridlines ws
    If ProductIsInspector() Then
        BuildStartHereInspector ws
    Else
        BuildStartHereStandard ws
    End If
    ws.Tab.Color = RGB(47, 79, 79)
End Sub

Private Sub BuildStartHereStandard(ByVal ws As Worksheet)
    TitleBlock ws, "RoadReviewer", _
        "Is it a federal-aid road? FHWA functional class + adjusted urban boundary checker."

    DisclaimerBlock ws, 5, 6

    ws.Range("B12").Value = "How to use:"
    ws.Range("B12").Font.Bold = True
    StepLine ws, 13, "1.  Pick your state below."
    StepLine ws, 14, "2.  Paste your Latitude and Longitude on the Sites tab (the yellow columns)."
    StepLine ws, 15, "3.  On the Sites tab, click Check Roads. Rows tint red (federal aid), green (non-federal aid) or yellow (review)."

    LabelValue ws, 17, "State", NR_STATE, "MI"
    LabelValue ws, 18, "Output Folder", NR_OUTFOLDER, ""
    LabelValue ws, 19, "AGOL Webmap URL (optional)", NR_AGOLMAP, ""
    LabelValue ws, 20, "Road/boundary search buffer (feet)", NR_BUFFER, CStr(DEFAULT_BUFFER_FEET)
    AddStateValidation ws.Cells(17, 3)
    SetOutputFolderDefault ws.Cells(18, 3)
    AddButton ws, ws.Cells(18, 4).Left + 6, ws.Cells(18, 4).Top - 2, 140, 22, "Browse for folder...", "SelectOutputFolder"
    AddBufferValidation ws.Cells(20, 3)

    NoteLine ws, 22, "The action buttons (Check Roads, Re-run, Photo Links, Export CSV/KML, Send to AGOL) are on the Sites tab's top row."
    NoteLine ws, 23, "Search buffer is how far to look for a road / urban boundary when the exact point misses. 250 ft is a good default."

    AddButton ws, 18, ws.Rows(26).Top, 170, 22, "Build / Reset Workbook", "BuildWorkbook", RGB(150, 150, 150)
    NoteLine ws, 28, "Output Folder shows where exports save (this workbook's folder); Browse or type to change it. " & _
        "Build / Reset repairs the layout - your Sites data is preserved."
    VersionLabel ws, 30
End Sub

Private Sub BuildStartHereInspector(ByVal ws As Worksheet)
    TitleBlock ws, "Site Inspector Review Tool", _
        "FEMA Public Assistance site inspection toolkit - classification, photos, FIRMettes and map pages."

    DisclaimerBlock ws, 5, 6

    StepLine ws, 12, "Fill in the job info, add points on the Sites tab, then run the numbered steps top to bottom."

    LabelValue ws, 14, "Work Order (WO #)", NR_WO, ""
    LabelValue ws, 15, "Impact (DI #)", NR_DI, ""
    LabelValue ws, 16, "Disaster Number", NR_DISASTER, ""
    LabelValue ws, 17, "Applicant", NR_APPLICANT, ""
    LabelValue ws, 18, "State", NR_STATE, "MI"
    LabelValue ws, 19, "Output Folder (optional)", NR_OUTFOLDER, ""
    LabelValue ws, 20, "AGOL Webmap URL (optional)", NR_AGOLMAP, ""
    LabelValue ws, 21, "Search buffer (feet)", NR_BUFFER, CStr(DEFAULT_BUFFER_FEET)
    AddStateValidation ws.Cells(18, 3)
    AddButton ws, ws.Cells(19, 4).Left + 6, ws.Cells(19, 4).Top - 2, 140, 22, "Browse for folder...", "SelectOutputFolder"
    AddBufferValidation ws.Cells(21, 3)

    NoteLine ws, 23, "WO/DI default onto every Sites row; override per row by typing in the row's WO/DI cells. " & _
        "Output Folder can stay blank - a OneDrive default is used and shown after each export."
    NoteLine ws, 24, "Search buffer is the fallback radius when no road intersects the exact point (and the floor for the " & _
        "urban-boundary check, min 250 ft). 250 ft is a good default; lower it for dense urban grids, raise it for sparse rural networks."
    NoteLine ws, 25, "MI / IN / WI road-class lookups are wired; other states still get the ACUB check. See the Sources tab for every layer + caveat."

    SectionLabel ws, 27, "1.  Classify Roads"
    AddButton ws, 18, ws.Rows(28).Top, 200, 32, "Check Roads", "CheckRoads", CLR_BTN_GO
    AddButton ws, 228, ws.Rows(28).Top, 170, 32, "Re-run Failed Rows", "ReRunFailedRows"

    SectionLabel ws, 31, "2.  Review Photos"
    AddButton ws, 18, ws.Rows(32).Top, 260, 30, "Open Photo Links for Selected Row(s)", "OpenImageryForSelection", CLR_BTN_GO

    SectionLabel ws, 35, "3.  Maps & FIRMettes"
    AddButton ws, 18, ws.Rows(36).Top, 190, 30, "Download FIRMettes", "DownloadFirmettes", CLR_BTN_GO
    AddButton ws, 218, ws.Rows(36).Top, 190, 30, "Re-run Failed FIRMettes", "ReRunFailedFirmettes"
    AddButton ws, 18, ws.Rows(39).Top, 190, 30, "Prepare Map Pages", "PrepareMapPages"
    AddButton ws, 218, ws.Rows(39).Top, 190, 30, "Add Blank Map Page", "AddMapPage"
    AddButton ws, 18, ws.Rows(42).Top, 250, 30, "Export Combined Map PDF", "ExportCombinedMapPdf"

    SectionLabel ws, 45, "Exports & hand-off"
    AddButton ws, 18, ws.Rows(46).Top, 190, 28, "Export Sites to KML", "ExportSitesToKML"
    AddButton ws, 218, ws.Rows(46).Top, 190, 28, "Export Sites Table (CSV)", "ExportSitesCsv"
    AddButton ws, 18, ws.Rows(49).Top, 390, 28, "Send Sites to AGOL Map (KML + open webmap)", "SendSitesToAgolMap"

    AddButton ws, 18, ws.Rows(52).Top, 170, 22, "Build / Reset Workbook", "BuildWorkbook", RGB(150, 150, 150)
    NoteLine ws, 54, "Build / Reset repairs the layout; your Sites data is preserved."
    VersionLabel ws, 56
End Sub

Private Sub AddStateValidation(ByVal cell As Range)
    With cell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=STATE_LIST
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

' Whole-number 1..1000 validation for the Search-buffer cell (both products).
' Out-of-range values still fall back to DEFAULT_BUFFER_FEET via
' modClassify.BufferFeet; this just catches typos at entry time.
Private Sub AddBufferValidation(ByVal cell As Range)
    With cell.Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, Formula1:="1", Formula2:="1000"
        .IgnoreBlank = True
        .ErrorMessage = "Enter a whole number between 1 and 1000 ft."
        .ShowError = True
    End With
End Sub

' Standard product only: show the folder this workbook is saved in right in
' the Output Folder cell, so the user sees where exports will land without
' having to configure anything. A live CELL("filename") formula yields the
' workbook's folder (with trailing "\"); it resolves to "" for an unsaved
' file or a web-hosted (SharePoint http://) path, which is exactly the case
' ResolveOutputFolder already falls back on - so behavior is unchanged, the
' cell just now displays the default. Browse-for-folder or typing overwrites
' the formula with an explicit path.
Private Sub SetOutputFolderDefault(ByVal cell As Range)
    cell.Formula = "=IF(OR(CELL(""filename"",$A$1)="""",LEFT(CELL(""filename"",$A$1),4)=""http""),""""," & _
        "LEFT(CELL(""filename"",$A$1),FIND(""["",CELL(""filename"",$A$1))-1))"
    cell.Font.Color = RGB(90, 90, 90)
    cell.Font.Italic = True
End Sub

Private Sub LabelValue(ByVal ws As Worksheet, ByVal r As Long, ByVal label As String, _
        ByVal namedRange As String, ByVal defaultVal As String)
    ws.Cells(r, 2).Value = label
    ws.Cells(r, 2).Font.Bold = True
    With ws.Cells(r, 3)
        .Interior.Color = CLR_INPUT
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        If Len(defaultVal) > 0 Then .Value = defaultVal
    End With
    AddNameForCell ws.Cells(r, 3), namedRange
End Sub

Private Sub AddNameForCell(ByVal cell As Range, ByVal nm As String)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nm, RefersTo:="='" & cell.Worksheet.Name & "'!" & cell.Address(True, True)
End Sub

' ---- Sites (the shared table) --------------------------------------------

Private Sub BuildSites()
    Dim ws As Worksheet
    Set ws = SheetByName(SH_SITES)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_SITES
    End If

    WriteSitesToolbar ws
    WriteSitesHeader ws
    FillSitesFormulas ws
    ApplySitesValidation ws
    ApplySitesFormatting ws
    ApplyProductColumns ws

    ws.Activate
    ws.Cells(SITES_FIRST_DATA_ROW, COL_SITENO).Select
    On Error Resume Next
    ' ActiveWindow can be Nothing when Excel is driven headless by COM.
    ' Skipping the freeze in that path is fine - the in-Excel "Build / Reset"
    ' button still pins the toolbar + header rows.
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    ws.Tab.Color = RGB(60, 60, 60)
End Sub

' Row 1 toolbar: the Sites actions live ON the Sites sheet so the paste ->
' classify -> review loop never leaves it. Buttons are free-floating so
' hiding columns (WO/DI, G-K) can't squash them. In the standard product
' EVERY common action is here (Start Here has no action buttons); the
' inspector keeps just the two shortcuts, its heavier workflow staying on
' its own Start Here toolbar.
Private Sub WriteSitesToolbar(ByVal ws As Worksheet)
    Dim i As Long
    ' Idempotent rebuild: drop only our own shapes (named RR_*).
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 3) = "RR_" Then ws.Shapes(i).Delete
    Next i

    ws.Rows(SITES_TOOLBAR_ROW).RowHeight = 26

    AddToolbarButton ws, "RR_CheckRoads", 2, 95, "Check Roads", "CheckRoads", True
    If ProductIsInspector() Then
        AddToolbarButton ws, "RR_PhotoLinks", 100, 190, "Photo Links (selected rows)", "OpenImageryForSelection", False
    Else
        AddToolbarButton ws, "RR_ReRun", 100, 105, "Re-run Failed", "ReRunFailedRows", False
        AddToolbarButton ws, "RR_PhotoLinks", 208, 150, "Photo Links (selected)", "OpenImageryForSelection", False
        AddToolbarButton ws, "RR_Csv", 361, 95, "Export CSV", "ExportSitesCsv", False
        AddToolbarButton ws, "RR_Kml", 459, 95, "Export KML", "ExportSitesToKML", False
        AddToolbarButton ws, "RR_Agol", 557, 175, "Send to AGOL Map", "SendSitesToAgolMap", False
    End If
End Sub

Private Sub AddToolbarButton(ByVal ws As Worksheet, ByVal shapeName As String, ByVal leftPt As Single, _
        ByVal widthPt As Single, ByVal caption As String, ByVal macroName As String, ByVal isGo As Boolean)
    Dim sh As Shape
    Set sh = AddButton(ws, leftPt, 2, widthPt, 22, caption, macroName, IIf(isGo, CLR_BTN_GO, CLR_BTN))
    sh.Name = shapeName
    sh.Placement = xlFreeFloating
End Sub

Private Sub WriteSitesHeader(ByVal ws As Worksheet)
    Dim h() As String
    ReDim h(1 To COL_LAST)
    h(COL_WO) = "WO #"
    h(COL_DI) = "DI #"
    h(COL_SITENO) = "Site #"
    h(COL_SITENAME) = "Site Name"
    h(COL_LAT) = "Latitude"
    h(COL_LON) = "Longitude"
    h(COL_DESC) = "Description (optional)"
    h(COL_ADDRESS) = "Address (optional)"
    h(COL_CATEGORY) = "Category (optional)"
    h(COL_COSTS) = "Costs (optional)"
    h(COL_WORKCOMP) = "Work Completion (optional)"
    h(COL_GEOCODE) = "Geocode Status"
    h(COL_GMAP) = "Google Maps"
    h(COL_STREETVIEW) = "Street View"
    h(COL_BING) = "Bing"
    h(COL_GEARTH) = "Google Earth"
    h(COL_FEMAVIEW) = "FEMA Viewer"
    h(COL_FIRMPORTAL) = "FIRMette Portal"
    h(COL_NFCMAP) = "NFC Map"
    h(COL_CLASS) = "FHWA Class"
    h(COL_URBANRURAL) = "Urban/Rural"
    h(COL_ACUBNAME) = "ACUB Name"
    h(COL_ROADNAME) = "Road Name"
    h(COL_STREET) = "Street Name"
    h(COL_ELIGIBILITY) = "Federal Aid Status"
    h(COL_REVIEWNOTE) = "Review Reason"
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
    SetLinkFormula ws, COL_GEARTH, r1, r2, URL_GEARTH, "Earth", latC, lonC
    SetLinkFormula ws, COL_FEMAVIEW, r1, r2, URL_FEMAVIEW, "FEMA", latC, lonC
    SetLinkFormula ws, COL_FIRMPORTAL, r1, r2, URL_FIRMPORTAL, "FIRMette", latC, lonC
    ' Per-row "Open in map" (F11), state-aware (F8): the URL depends on
    ' Start Here's State dropdown, not just lat/lon, so it gets its own
    ' formula builder rather than SetLinkFormula's single-template pattern.
    SetNfcMapFormula ws, r1, r2, latC, lonC
    ' The AGOL Map column is driven by the user's own webmap URL on Start
    ' Here. The formula handles all three "blank" states (no URL set,
    ' missing coords) and stitches center/level/marker query params onto
    ' whatever URL was pasted.
    SetAgolMapFormula ws, r1, r2, latC, lonC
End Sub

' The AGOL formula is bespoke enough (depends on the dynamic Start Here URL,
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
    urlExpr = UrlExprFromTemplate(urlTemplate, latC, lonC)
    If needsCoords Then
        f = "=IF(OR(" & latC & "="""" ," & lonC & "=""""),"""",HYPERLINK(" & urlExpr & ",""" & friendly & """))"
    Else
        f = "=HYPERLINK(" & urlExpr & ",""" & friendly & """)"
    End If
    ws.Range(ws.Cells(r1, col), ws.Cells(r2, col)).Formula = f
End Sub

' Builds an Excel formula-EXPRESSION (not a full formula - no leading "=")
' for a URL template with {LAT}/{LON} placeholders substituted by cell
' references, e.g. """https://x?lat=""&$F2&""&lon=""&$G2".
Private Function UrlExprFromTemplate(ByVal urlTemplate As String, ByVal latC As String, ByVal lonC As String) As String
    UrlExprFromTemplate = """" & Replace(Replace(urlTemplate, "{LAT}", """&" & latC & "&"""), "{LON}", """&" & lonC & "&""") & """"
End Function

' Per-row "Open in map" (F11), keyed off Start Here's State dropdown (F8)
' since MI/IN/WI each need a different URL (§4.2/§4.2a/§4.2b) - the only
' wired states get their own map link; anything else (MN/IL/OH, or blank)
' falls back to the plain FEMA Map Viewer pin so the column is never broken,
' just generic.
Private Sub SetNfcMapFormula(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
        ByVal latC As String, ByVal lonC As String)
    Dim miExpr As String, inExpr As String, wiExpr As String, fallbackExpr As String
    Dim urlExpr As String, f As String
    miExpr = UrlExprFromTemplate(URL_NFC_MAPVIEW, latC, lonC)
    inExpr = UrlExprFromTemplate(URL_NFC_MAPVIEW_IN, latC, lonC)
    wiExpr = UrlExprFromTemplate(URL_NFC_MAPVIEW_WI, latC, lonC)
    fallbackExpr = UrlExprFromTemplate(URL_FEMAVIEW, latC, lonC)
    ' Blank State (never happens once Start Here is built - the State cell
    ' defaults to "MI" - but cheap to guard) matches ClassifyRows's
    ' default-to-MI behavior.
    urlExpr = "IF(OR(" & NR_STATE & "=""MI""," & NR_STATE & "="""")," & miExpr & _
        ",IF(" & NR_STATE & "=""IN""," & inExpr & _
        ",IF(" & NR_STATE & "=""WI""," & wiExpr & "," & fallbackExpr & ")))"
    f = "=IF(OR(" & latC & "="""" ," & lonC & "=""""),"""",HYPERLINK(" & urlExpr & ",""Open""))"
    ws.Range(ws.Cells(r1, COL_NFCMAP), ws.Cells(r2, COL_NFCMAP)).Formula = f
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
    ws.Columns(COL_DESC).ColumnWidth = 26
    ws.Columns(COL_ADDRESS).ColumnWidth = 28
    ws.Columns(COL_CATEGORY).ColumnWidth = 12
    ws.Columns(COL_COSTS).ColumnWidth = 14
    ws.Columns(COL_WORKCOMP).ColumnWidth = 16
    ws.Columns(COL_CLASS).ColumnWidth = 16
    ws.Columns(COL_ACUBNAME).ColumnWidth = 16
    ws.Columns(COL_ROADNAME).ColumnWidth = 16
    ws.Columns(COL_STREET).ColumnWidth = 22
    ws.Columns(COL_ELIGIBILITY).ColumnWidth = 24
    ws.Columns(COL_REVIEWNOTE).ColumnWidth = 18
    ws.Columns(COL_GEOCODE).ColumnWidth = 16
    ws.Columns(COL_FIRMSTATUS).ColumnWidth = 18
    ws.Columns(COL_MAPSTATUS).ColumnWidth = 16
    ws.Columns(COL_AGOLMAP).ColumnWidth = 10

    ' Guidance tints (friction fix: nothing used to tell the user which
    ' columns are theirs). Yellow = type here; grey = a workflow writes it.
    ' Link columns stay white. Conditional-format rules below win over
    ' these static fills when they match, so the verdict colors still show.
    ws.Range(ws.Cells(SITES_FIRST_DATA_ROW, COL_WO), ws.Cells(r2, COL_WORKCOMP)).Interior.Color = CLR_INPUT
    ws.Range(ws.Cells(SITES_FIRST_DATA_ROW, COL_GEOCODE), ws.Cells(r2, COL_GEOCODE)).Interior.Color = CLR_RESULT
    ws.Range(ws.Cells(SITES_FIRST_DATA_ROW, COL_CLASS), ws.Cells(r2, COL_REVIEWNOTE)).Interior.Color = CLR_RESULT
    ws.Range(ws.Cells(SITES_FIRST_DATA_ROW, COL_FIRMSTATUS), ws.Cells(r2, COL_MAPSTATUS)).Interior.Color = CLR_RESULT

    ' Tri-state highlight on the Federal Aid Status column:
    '   red    — cell starts with "Federal aid"  (federal-aid road)
    '   green  — cell starts with "Non-federal aid"
    '   yellow — cell starts with "Review" (non-certified class or no
    '            road found within the search-buffer radius)
    ' "Non-federal aid" intentionally tests for the literal prefix
    ' (LEFT … 15) because a substring search for "Federal aid" would
    ' also match "Non-federal aid". Order matters when format rules
    ' overlap: in Excel the FIRST matching rule wins, so non-federal
    ' is checked before federal.
    eligCol = "$" & ColLetter(COL_ELIGIBILITY) & SITES_FIRST_DATA_ROW
    With ws.Range(ws.Cells(SITES_FIRST_DATA_ROW, COL_CLASS), ws.Cells(r2, COL_REVIEWNOTE))
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
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(SITES_HEADER_ROW, 1), ws.Cells(SITES_HEADER_ROW, COL_LAST)).AutoFilter
    On Error GoTo 0
End Sub

' The inspector-only columns exist in both products (shared COL_* constants,
' shared workflow code) but are hidden in the standard product so a PDMG /
' partner / reviewer never sees job bookkeeping they don't use.
Private Sub ApplyProductColumns(ByVal ws As Worksheet)
    Dim hideCols As Boolean
    hideCols = Not ProductIsInspector()
    ws.Columns(COL_WO).Hidden = hideCols
    ws.Columns(COL_DI).Hidden = hideCols
    ws.Columns(COL_FIRMSTATUS).Hidden = hideCols
    ws.Columns(COL_MAPSTATUS).Hidden = hideCols

    ' Optional user-data columns (Description..Work Completion = cols G-K) are
    ' hidden by default on both products to keep the paste area tight around
    ' Latitude/Longitude. Unhide to enter them; their values still flow into
    ' the KML/CSV exports and (inspector) the map-page stamps when present.
    Dim c As Long
    For c = COL_DESC To COL_WORKCOMP
        ws.Columns(c).Hidden = True
    Next c
End Sub
