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

' Left edge (points, from the sheet's left) for anything that sits UNDER a
' section heading: buttons and dropdowns. Matches the one-level cell indent
' StepLine/NoteLine apply, so a section's text and its buttons share a margin
' while the heading itself stays flush left.
Private Const BODY_LEFT_PT As Single = 30

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

' "Reset Everything" button. Unlike BuildWorkbook (which preserves typed Sites
' data), this deletes the Sites and MapPages sheets outright so the rebuild
' starts from a blank table - the fix for a Sites sheet that got into a bad
' state. Destructive, so it confirms first (skipped only under headless
' automation).
Public Sub ResetWorkbookFull()
    If Not gHeadless Then
        If MsgBox("This ERASES every point on the Sites sheet and rebuilds a blank workbook." & vbCrLf & vbCrLf & _
            "This cannot be undone. Continue?", _
            vbExclamation + vbYesNo + vbDefaultButton2, "Reset Everything") <> vbYes Then Exit Sub
    End If

    Application.DisplayAlerts = False
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = SheetByName(SH_SITES)
    If Not ws Is Nothing Then ws.Delete
    Set ws = SheetByName(SH_MAPPAGES)
    If Not ws Is Nothing Then ws.Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    BuildWorkbook          ' recreates a blank Sites table + relays its own result dialog
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
    ' Inspector: hide the Sources tab (it stays in the file - still powers the
    ' citations echo and service-URL overrides - just isn't shown to the
    ' inspector). Activate a visible sheet first so we're never hiding the
    ' active sheet, which Excel rejects.
    Dim wsSrc As Worksheet
    Set wsSrc = SheetByName(SH_SOURCES)
    If wsSrc Is Nothing Then Exit Sub
    On Error Resume Next
    If ProductIsInspector() Then
        SheetByName(SH_START).Activate
        wsSrc.Visible = xlSheetHidden
    Else
        wsSrc.Visible = xlSheetVisible
    End If
    On Error GoTo 0
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

' Body copy under a section heading. Indented one level so it reads as
' belonging to the heading above it (BODY_INDENT_PT keeps the buttons on the
' same left edge as this text).
Private Sub NoteLine(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    With ws.Cells(r, 2)
        .Value = txt
        .Font.Italic = True
        .Font.Color = RGB(90, 90, 90)
        .IndentLevel = 1
    End With
End Sub

Private Sub StepLine(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    With ws.Cells(r, 2)
        .Value = txt
        .Font.Color = RGB(70, 70, 70)
        .IndentLevel = 1
    End With
End Sub

' Section heading: larger, bold, in the workbook's slate accent, with a rule
' under it spanning B:C. Sits flush left (its body indents beneath it) so the
' page reads as headed sections rather than one flat column of bold text.
Private Sub SectionLabel(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    With ws.Cells(r, 2)
        .Value = txt
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(47, 79, 79)
        .IndentLevel = 0
    End With
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(190, 200, 205)
        .Weight = xlThin
    End With
    ws.Rows(r).RowHeight = 22
End Sub

' Shared disclaimer wording. Two surfaces use it: the standard product's
' on-sheet red box (DisclaimerBlock below) and the inspector product's dialog,
' which modClassify.CheckRoads pops once per session instead of carrying the
' box on Start Here (per user request). Exposed as functions so both surfaces
' stay byte-identical. The old "It is not an authoritative source for FHWA
' functional classification." sentence was dropped per user request.
Public Function DisclaimerHeaderText() As String
    DisclaimerHeaderText = "IMPORTANT - NOT AN AUTHORITATIVE FHWA OR ELIGIBILITY DETERMINATION"
End Function

Public Function DisclaimerBodyText() As String
    DisclaimerBodyText = "This tool does NOT authoritatively identify FHWA federal-aid roads. It flags high-probability " & _
        "candidates for a person to review, and may miss or mis-tag roads. EVERY coordinate must be verified by a human " & _
        "against the official source map - use each row's NFC Map link and the Sources tab. Results are informational " & _
        "only."
End Function

' Prominent red-bordered disclaimer box spanning B:C over `rowCount` rows.
' Standard product only now - the inspector shows the same text as a dialog on
' Check Roads. Kept in sync with modSources' echo of it and web/index.html.
Private Sub DisclaimerBlock(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal rowCount As Long)
    Dim rng As Range, r As Long, hdr As String, body As String
    hdr = DisclaimerHeaderText()
    body = DisclaimerBodyText()
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
    TitleBlock ws, "RoadReviewer", "Federal-aid road checker and review tool"

    ' 4 rows, not 6 - the disclaimer body was trimmed (ends at "informational
    ' only"), so the box shrinks to fit rather than carrying dead space.
    DisclaimerBlock ws, 5, 4

    SectionLabel ws, 12, "How To Use"
    StepLine ws, 13, "1.  Pick your state below."
    StepLine ws, 14, "2.  Paste your Latitude and Longitude on the Sites tab (the yellow columns)."
    StepLine ws, 15, "3.  Come back here and click Check Roads. Rows tint red (federal aid), green (non-federal aid) or yellow (review)."

    LabelValue ws, 17, "State", NR_STATE, "MI"
    LabelValue ws, 18, "Output Folder", NR_OUTFOLDER, ""
    LabelValue ws, 19, "User-Defined AGOL Layer", NR_AGOLMAP, ""
    ' Same label as the inspector's - the old "Road/boundary search buffer
    ' (feet)" overran column B and got clipped by the yellow input cell.
    LabelValue ws, 20, "FHWA search buffer (feet)", NR_BUFFER, CStr(DEFAULT_BUFFER_FEET)
    AddStateValidation ws.Cells(17, 3)
    SetOutputFolderDefault ws.Cells(18, 3)
    AddButton ws, ws.Cells(18, 4).Left + 6, ws.Cells(18, 4).Top - 2, 140, 22, "Browse for folder...", "SelectOutputFolder"
    AddBufferValidation ws.Cells(20, 3)

    NoteLine ws, 23, "Search buffer is how far to look for a road / urban boundary when the exact point misses."

    ' Actions live here now, not on a Sites row-1 toolbar (row 1 is the header
    ' row). Same macros, same order the old toolbar used. The "Photo Links
    ' (selected rows)" button was dropped per user request; OpenImageryForSelection
    ' is still a live macro (and still reachable from the inspector's roads
    ' dropdown), so the button can be restored here without touching modImagery.
    ' Section titles mirror the inspector's (no parentheticals) - the standard
    ' product has no imagery button, so its check section is FHWA-status only.
    SectionLabel ws, 25, "Check FHWA Status"
    AddButton ws, BODY_LEFT_PT, ws.Rows(26).Top, 200, 30, "Check Roads", "CheckRoads", CLR_BTN_GO
    AddButton ws, BODY_LEFT_PT + 210, ws.Rows(26).Top, 170, 30, "Re-run Failed Rows", "ReRunFailedRows"

    SectionLabel ws, 29, "Exports & Handoff"
    CreateExportPicker ws, BODY_LEFT_PT, ws.Rows(30).Top + 3, 300
    AddButton ws, BODY_LEFT_PT + 308, ws.Rows(30).Top, 70, 24, "Go", "RunSelectedExport", CLR_BTN_GO
    NoteLine ws, 32, "Pick an export, then click Go. FIRMettes and map PDFs save to the Output Folder above."

    SectionLabel ws, 34, "Repair / Reset"
    AddButton ws, BODY_LEFT_PT, ws.Rows(35).Top, 210, 24, "Repair Layout (keeps your data)", "BuildWorkbook", RGB(120, 120, 120)
    AddButton ws, BODY_LEFT_PT + 220, ws.Rows(35).Top, 220, 24, "Reset Everything (erases data)", "ResetWorkbookFull", RGB(176, 80, 80)
    NoteLine ws, 37, "Repair Layout rebuilds the sheets, buttons and formulas and KEEPS your typed Sites data. " & _
        "Reset Everything deletes every point and rebuilds a blank Sites table - it asks you to confirm first."
    VersionLabel ws, 39
End Sub

Private Sub BuildStartHereInspector(ByVal ws As Worksheet)
    ' No subtitle row (row 3 dropped per user request) and no on-sheet
    ' disclaimer box - the not-authoritative disclaimer is shown as a dialog
    ' the first time Check Roads runs (modClassify.ShowDisclaimerOnce).
    With ws.Range("B2")
        .Value = "Site Inspector Review Tool"
        .Font.Size = 20
        .Font.Bold = True
    End With

    StepLine ws, 4, "Fill in the job info, add points on the Sites tab, then work top to bottom."

    LabelValue ws, 6, "Work Order (WO #)", NR_WO, ""
    LabelValue ws, 7, "Impact (DI #)", NR_DI, ""
    LabelValue ws, 8, "Disaster Number", NR_DISASTER, ""
    LabelValue ws, 9, "Applicant", NR_APPLICANT, ""
    LabelValue ws, 10, "State", NR_STATE, "MI"
    LabelValue ws, 11, "Output Folder (optional)", NR_OUTFOLDER, ""
    LabelValue ws, 12, "User-Defined AGOL Layer", NR_AGOLMAP, ""
    LabelValue ws, 13, "FHWA search buffer (feet)", NR_BUFFER, CStr(DEFAULT_BUFFER_FEET)
    AddStateValidation ws.Cells(10, 3)
    AddButton ws, ws.Cells(11, 4).Left + 6, ws.Cells(11, 4).Top - 2, 140, 22, "Browse for folder...", "SelectOutputFolder"
    AddBufferValidation ws.Cells(13, 3)

    NoteLine ws, 15, "WO/DI default onto every Sites row; override per row by typing in the row's WO/DI cells. " & _
        "Output Folder can stay blank - exports save next to this workbook."
    NoteLine ws, 16, "FHWA search buffer is the fallback radius when no road intersects the exact point (and the floor for the " & _
        "urban-boundary check, min 250 ft). 250 ft is a good default; lower it for dense grids, raise it for sparse rural networks."
    NoteLine ws, 17, "MI / IN / WI road-class lookups are wired; other states still get the ACUB check."

    ' ---- Location Maps ----
    ' KML + screenshots come BEFORE Prepare Map Pages (steps 2/3 swapped per
    ' user request): gather the images first, then build the pages that step 4
    ' drops them onto. PrepareMapPages doesn't read the images, so either order
    ' works - this one just matches how the job actually goes.
    SectionLabel ws, 19, "Location Maps"
    StepLine ws, 21, "1.  Input site info on the Sites tab - coordinates, site name, WO/DI and category."
    StepLine ws, 23, "2.  Export the sites to KML, open it (Google Earth), and screenshot each site."
    AddButton ws, BODY_LEFT_PT, ws.Rows(24).Top, 190, 30, "Export Sites to KML", "ExportSitesToKML", CLR_BTN_GO
    NoteLine ws, 26, "Press Windows+Shift+S, drag a box around the site, and save each image in a 'maps' subfolder of the " & _
        "Output Folder - or use each page's 'Select photo' button."
    StepLine ws, 28, "3.  Prepare the map pages - one landscape page per site (opens the MapPages tab)."
    AddButton ws, BODY_LEFT_PT, ws.Rows(29).Top, 190, 30, "Prepare Map Pages", "PrepareMapPages", CLR_BTN_GO
    AddButton ws, BODY_LEFT_PT + 200, ws.Rows(29).Top, 190, 30, "Add Blank Map Page", "AddMapPage"
    StepLine ws, 32, "4.  On the MapPages tab (right of the pages) use 'Map page tools': click 'Insert Map Images' to drop your " & _
        "screenshots on, oldest first - or a page's 'Select photo' button to place one by hand."
    ' Step 5's long explanation was dropped per user request - the button says
    ' what it does. The same button also lives on the MapPages "Map page tools"
    ' panel, so the inspector can export from either sheet.
    StepLine ws, 34, "5.  Export the combined map PDF."
    AddButton ws, BODY_LEFT_PT, ws.Rows(35).Top, 230, 30, "Export Combined Map PDF", "ExportCombinedMapPdf", CLR_BTN_GO

    ' ---- FIRMette ----
    SectionLabel ws, 39, "FIRMette"
    AddButton ws, BODY_LEFT_PT, ws.Rows(40).Top, 190, 30, "Download FIRMettes", "DownloadFirmettes", CLR_BTN_GO
    AddButton ws, BODY_LEFT_PT + 200, ws.Rows(40).Top, 210, 30, "Re-run Failed FIRMettes", "ReRunFailedFirmettes"
    NoteLine ws, 42, "Downloads the FEMA FIRMette PDF for every site to the Output Folder. Re-run only retries failed rows."

    ' ---- Check FHWA Status and Imagery: one dropdown (classify / re-run / photo tabs) ----
    SectionLabel ws, 45, "Check FHWA Status and Imagery"
    CreateRoadsPicker ws, BODY_LEFT_PT, ws.Rows(46).Top + 3, 300
    AddButton ws, BODY_LEFT_PT + 308, ws.Rows(46).Top, 70, 24, "Go", "RunSelectedRoadsAction", CLR_BTN_GO
    NoteLine ws, 48, "Pick an action, then click Go: classify roads, re-run failed rows, or open photo tabs for the selected row(s)."

    ' ---- Exports & Handoff ----
    SectionLabel ws, 51, "Exports & Handoff"
    CreateExportPicker ws, BODY_LEFT_PT, ws.Rows(52).Top + 3, 300
    AddButton ws, BODY_LEFT_PT + 308, ws.Rows(52).Top, 70, 24, "Go", "RunSelectedExport", CLR_BTN_GO
    NoteLine ws, 54, "Pick an export, then click Go. Everything writes to the Output Folder above."

    SectionLabel ws, 57, "Repair / Reset"
    AddButton ws, BODY_LEFT_PT, ws.Rows(58).Top, 210, 24, "Repair Layout (keeps your data)", "BuildWorkbook", RGB(120, 120, 120)
    AddButton ws, BODY_LEFT_PT + 220, ws.Rows(58).Top, 220, 24, "Reset Everything (erases data)", "ResetWorkbookFull", RGB(176, 80, 80)
    NoteLine ws, 60, "Repair Layout rebuilds the sheets, buttons and formulas and KEEPS your typed Sites data. " & _
        "Reset Everything deletes every point and rebuilds a blank Sites table - it asks you to confirm first."
    VersionLabel ws, 62
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
        ' Right-justified per user request: the buffer cell was already
        ' right-reading (it's a number) and the text inputs now line up with it.
        .HorizontalAlignment = xlRight
        If Len(defaultVal) > 0 Then .Value = defaultVal
    End With
    AddNameForCell ws.Cells(r, 3), namedRange
End Sub

Public Sub AddNameForCell(ByVal cell As Range, ByVal nm As String)
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

' The Sites sheet no longer carries a toolbar - row 1 is the header row now,
' and the old buttons were free-floating shapes pinned at Top:=2pt, so they
' sat directly over it. Every action moved back to Start Here. This sub is
' kept (and still called) purely so a Build / Reset on a workbook built by an
' older version strips the orphaned RR_* shapes instead of leaving them
' floating over the headers.
Private Sub WriteSitesToolbar(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 3) = "RR_" Then ws.Shapes(i).Delete
    Next i
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
    h(COL_NFCAGOL) = "NFC Layer (Map Viewer)"
    h(COL_AGOLMAP) = "User-Defined AGOL Layer"
    h(COL_NFCMAP) = "State NFC App"
    h(COL_CLASS) = "FHWA Class"
    h(COL_URBANRURAL) = "Urban/Rural"
    h(COL_ACUBNAME) = "ACUB Name"
    h(COL_ROADNAME) = "Road Name"
    h(COL_STREET) = "Street Name"
    h(COL_ELIGIBILITY) = "Federal Aid Status"
    h(COL_REVIEWNOTE) = "Review Reason"
    h(COL_FIRMSTATUS) = "FIRMette Status"
    h(COL_MAPSTATUS) = "Map Status"

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

    SetSiteNumberFormula ws, r1, r2
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
    ' AGOL NFC Layer column: the state functional-class layer in ArcGIS Map
    ' Viewer, centered on the point (MI = curated webmap to avoid the time
    ' slider; IN/WI = live side-load; others = plain FEMA pin).
    SetNfcAgolFormula ws, r1, r2, latC, lonC
End Sub

' Site # is auto-numbered by formula (no more "type a site number" step):
' every row that has a Latitude gets a running 1, 2, 3 ... while blank rows
' stay blank. Assigning one formula to the whole column lets Excel adjust the
' relative refs per row (the COUNT window's start stays absolute at row r1).
' Typing a value into a Site # cell still overrides the formula for that row.
Private Sub SetSiteNumberFormula(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long)
    Dim latAbs As String, latRel As String, f As String
    latAbs = "$" & ColLetter(COL_LAT) & "$" & r1     ' e.g. $E$2 (fixed window start)
    latRel = "$" & ColLetter(COL_LAT) & r1           ' e.g. $E2  (row-relative)
    f = "=IF(" & latRel & "="""","""",COUNT(" & latAbs & ":" & latRel & "))"
    ws.Range(ws.Cells(r1, COL_SITENO), ws.Cells(r2, COL_SITENO)).Formula = f
End Sub

' An Excel string literal "s" - used when a full URL (not a {LAT}/{LON}
' template) is dropped into a formula, so the quoting stays legible.
Private Function ExcelStr(ByVal s As String) As String
    ExcelStr = Chr$(34) & s & Chr$(34)
End Function

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

' NFC Map / "Open" column (COL_NFCMAP): the state's official PUBLIC APP,
' keyed off the State dropdown. App URLs carry no coordinates (the Experience
' apps can't be reliably centered via URL - PR #17/#18), so a row just opens
' the authoritative app; the AGOL NFC Layer column is the one that centers on
' the exact point. Blank State defaults to MI (matches ClassifyRows).
Private Sub SetNfcMapFormula(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
        ByVal latC As String, ByVal lonC As String)
    Dim urlExpr As String, f As String
    urlExpr = "IF(OR(" & NR_STATE & "=" & ExcelStr("MI") & "," & NR_STATE & "=" & ExcelStr("") & ")," & ExcelStr(APP_MI) & _
        ",IF(" & NR_STATE & "=" & ExcelStr("IN") & "," & ExcelStr(APP_IN) & _
        ",IF(" & NR_STATE & "=" & ExcelStr("WI") & "," & ExcelStr(APP_WI) & _
        ",IF(" & NR_STATE & "=" & ExcelStr("MN") & "," & ExcelStr(APP_MN) & _
        ",IF(" & NR_STATE & "=" & ExcelStr("IL") & "," & ExcelStr(APP_IL) & _
        ",IF(" & NR_STATE & "=" & ExcelStr("OH") & "," & ExcelStr(APP_OH) & "," & ExcelStr(APP_MI) & "))))))"
    f = "=IF(OR(" & latC & "="""" ," & lonC & "=""""),"""",HYPERLINK(" & urlExpr & ",""Open""))"
    ws.Range(ws.Cells(r1, COL_NFCMAP), ws.Cells(r2, COL_NFCMAP)).Formula = f
End Sub

' AGOL NFC Layer column (COL_NFCAGOL): the state functional-class layer in
' ArcGIS Map Viewer, centered + markered on the row's point. This is the old
' per-state Map Viewer link (MI now uses the curated webmap so no time slider;
' IN/WI side-load their live layer; others get the plain FEMA pin).
Private Sub SetNfcAgolFormula(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
        ByVal latC As String, ByVal lonC As String)
    Dim miExpr As String, inExpr As String, wiExpr As String, fallbackExpr As String
    Dim urlExpr As String, f As String
    miExpr = UrlExprFromTemplate(URL_NFC_MAPVIEW, latC, lonC)
    inExpr = UrlExprFromTemplate(URL_NFC_MAPVIEW_IN, latC, lonC)
    wiExpr = UrlExprFromTemplate(URL_NFC_MAPVIEW_WI, latC, lonC)
    fallbackExpr = UrlExprFromTemplate(URL_FEMAVIEW, latC, lonC)
    urlExpr = "IF(OR(" & NR_STATE & "=""MI""," & NR_STATE & "="""")," & miExpr & _
        ",IF(" & NR_STATE & "=""IN""," & inExpr & _
        ",IF(" & NR_STATE & "=""WI""," & wiExpr & "," & fallbackExpr & ")))"
    f = "=IF(OR(" & latC & "="""" ," & lonC & "=""""),"""",HYPERLINK(" & urlExpr & ",""Open""))"
    ws.Range(ws.Cells(r1, COL_NFCAGOL), ws.Cells(r2, COL_NFCAGOL)).Formula = f
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
    ws.Columns(COL_NFCAGOL).ColumnWidth = 14

    ' Guidance tints (friction fix: nothing used to tell the user which
    ' columns are theirs). Yellow = type here; grey = a workflow writes it.
    ' Link columns stay white. Conditional-format rules below win over
    ' these static fills when they match, so the verdict colors still show.
    ws.Range(ws.Cells(SITES_FIRST_DATA_ROW, COL_WO), ws.Cells(r2, COL_WORKCOMP)).Interior.Color = CLR_INPUT
    ' Site # is a formula now (auto-numbered), not a typed input: tint it grey
    ' like the other computed columns and center it so it reads as generated.
    With ws.Range(ws.Cells(SITES_FIRST_DATA_ROW, COL_SITENO), ws.Cells(r2, COL_SITENO))
        .Interior.Color = CLR_RESULT
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(SITES_FIRST_DATA_ROW, COL_GEOCODE), ws.Cells(r2, COL_GEOCODE)).Interior.Color = CLR_RESULT
    ws.Range(ws.Cells(SITES_FIRST_DATA_ROW, COL_CLASS), ws.Cells(r2, COL_REVIEWNOTE)).Interior.Color = CLR_RESULT
    ws.Range(ws.Cells(SITES_FIRST_DATA_ROW, COL_FIRMSTATUS), ws.Cells(r2, COL_MAPSTATUS)).Interior.Color = CLR_RESULT

    ' Make the HYPERLINK()-formula columns LOOK like links (blue + underline);
    ' Excel doesn't auto-style formula-driven hyperlinks the way it does
    ' clicked-in ones. Empty cells just carry the style with no visible text.
    Dim lc As Variant
    For Each lc In Array(COL_NFCMAP, COL_GMAP, COL_STREETVIEW, COL_BING, COL_GEARTH, COL_FEMAVIEW, COL_FIRMPORTAL, COL_AGOLMAP, COL_NFCAGOL)
        With ws.Range(ws.Cells(SITES_FIRST_DATA_ROW, CLng(lc)), ws.Cells(r2, CLng(lc))).Font
            .Color = RGB(5, 99, 193)
            .Underline = xlUnderlineStyleSingle
        End With
    Next lc

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

    ' Optional user-data columns. Description (G), Address (H) and Costs (J)
    ' stay hidden to keep the paste area tight around Latitude/Longitude;
    ' Category (I) and Work Completion (K) are shown per user request. Hidden
    ' values still flow into the KML/CSV exports and (inspector) the map-page
    ' stamps when present.
    ws.Columns(COL_DESC).Hidden = True        ' G
    ws.Columns(COL_ADDRESS).Hidden = True     ' H
    ws.Columns(COL_CATEGORY).Hidden = False   ' I - shown
    ws.Columns(COL_COSTS).Hidden = True       ' J
    ws.Columns(COL_WORKCOMP).Hidden = False   ' K - shown
    ' Geocode Status (L) is hidden per user request - geocode failures also
    ' surface in the Federal Aid Status column, so nothing is lost by hiding it.
    ws.Columns(COL_GEOCODE).Hidden = True      ' L

    ' Inspector only: Bing and Google Earth stay hidden for good (the inspector
    ' works from Google Maps / Street View / FEMA). The standard product keeps
    ' its full photo-link strip.
    If ProductIsInspector() Then
        ws.Columns(COL_BING).Hidden = True
        ws.Columns(COL_GEARTH).Hidden = True
    Else
        ws.Columns(COL_BING).Hidden = False
        ws.Columns(COL_GEARTH).Hidden = False
    End If

    ' Auto-reviewer output columns start hidden (and, on the inspector, so do
    ' the two NFC map-link columns - they only mean something once a row has a
    ' class). CheckRoads / ReRunFailedRows reveal them once there is something
    ' to show. A Build / Reset on a sheet that already holds results keeps them
    ' visible rather than hiding data.
    If SitesHasClassifiedRows(ws) Then
        ShowReviewerColumns
    Else
        HideReviewerColumns
    End If
End Sub

' ---- auto-reviewer column visibility --------------------------------------
'
' COL_REVIEWER_FIRST..COL_REVIEWER_LAST (FHWA Class .. Review Reason) are
' written ONLY by the classifier. Showing them on a blank sheet advertises
' seven empty columns and pushes the imagery links off-screen, so they stay
' hidden until the macro that fills them has run.

Public Sub ShowReviewerColumns()
    SetReviewerColumnsHidden False
End Sub

Public Sub HideReviewerColumns()
    SetReviewerColumnsHidden True
End Sub

Private Sub SetReviewerColumnsHidden(ByVal hide As Boolean)
    On Error Resume Next        ' never let cosmetics break a classify run
    Dim ws As Worksheet, c As Long
    Set ws = SitesSheet()
    If ws Is Nothing Then Exit Sub
    For c = COL_REVIEWER_FIRST To COL_REVIEWER_LAST
        ws.Columns(c).Hidden = hide
    Next c
    ' Inspector only: the two NFC map links ride along with the FHWA results -
    ' they're the "go look at the source map" follow-up to a class, so they show
    ' up exactly when the class does. The standard product shows them always.
    If ProductIsInspector() Then
        ws.Columns(COL_NFCAGOL).Hidden = hide
        ws.Columns(COL_NFCMAP).Hidden = hide
    End If
    On Error GoTo 0
End Sub

' True when any data row already carries a verdict - i.e. the classifier has
' run at some point and its columns should not be re-hidden by a rebuild.
Private Function SitesHasClassifiedRows(ByVal ws As Worksheet) As Boolean
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_LAT).End(xlUp).Row
    If lastRow < SITES_FIRST_DATA_ROW Then Exit Function

    For r = SITES_FIRST_DATA_ROW To lastRow
        If Len(Trim$(CStr(ws.Cells(r, COL_ELIGIBILITY).Value))) > 0 Then
            SitesHasClassifiedRows = True
            Exit Function
        End If
    Next r
End Function
