Attribute VB_Name = "modUtil"
Option Explicit

' RoadReviewer V1 - shared helpers used across workflows.
' Every sheet reference is hard-bound by name (friction point #4): no
' workflow ever writes to ActiveSheet.

' ---- Module-level state ---------------------------------------------------
' All Public variables must live in the declarations section (before the
' first Sub/Function) or VBA throws a compile error.

' Set True by automation hosts (build\build.ps1, build\verify-*.ps1) to
' suppress user-facing MsgBox prompts so the COM caller doesn't hang on
' an invisible modal dialog. Workflows still write all status to cells +
' StatusBar, so the result is fully observable when this flag is True.
' Defaults to False so the in-Excel button-click experience is unchanged.
Public gHeadless As Boolean

' Diagnostic trace file path. Off by default; flipped on by verify-*.ps1
' so HTTP calls + per-row checkpoints get appended to a known file. When a
' long-running workflow hangs, the trace tells us exactly where it stopped
' instead of leaving us staring at a frozen Excel window.
Public gTracePath As String

' ---- Setters callable from a COM automation host -------------------------
' Application.Run can only invoke named subs, not assign module-level
' variables directly, so the host calls Application.Run "SetHeadless"/"SetTrace".

Public Sub SetHeadless(ByVal value As Boolean)
    gHeadless = value
End Sub

Public Sub SetTrace(ByVal path As String)
    gTracePath = path
    On Error Resume Next
    Kill path
    On Error GoTo 0
End Sub

' ---- Product identity ------------------------------------------------------
' Which of the two products this workbook is (PRODUCT_STANDARD /
' PRODUCT_INSPECTOR) is baked in at build time as a hidden defined name, so
' it survives save/reopen and the in-Excel "Build / Reset Workbook" button
' rebuilds the same product. build.ps1 calls SetProduct via Application.Run
' before running BuildWorkbook.

Public Sub SetProduct(ByVal productName As String)
    On Error Resume Next
    ThisWorkbook.Names(NM_PRODUCT).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=NM_PRODUCT, RefersTo:="=""" & productName & """", Visible:=False
End Sub

' Missing/garbled name defaults to Inspector - the superset product - so a
' workbook built before this flag existed keeps its full behavior.
Public Function ProductName() As String
    Dim v As String
    On Error Resume Next
    v = ThisWorkbook.Names(NM_PRODUCT).RefersTo   ' looks like ="Standard"
    On Error GoTo 0
    v = Replace(Replace(v, "=", ""), """", "")
    If v <> PRODUCT_STANDARD Then v = PRODUCT_INSPECTOR
    ProductName = v
End Function

Public Function ProductIsInspector() As Boolean
    ProductIsInspector = (ProductName() = PRODUCT_INSPECTOR)
End Function

Public Function ProductTitle() As String
    If ProductIsInspector() Then
        ProductTitle = "Site Inspector Review Tool"
    Else
        ProductTitle = "RoadReviewer"
    End If
End Function

Public Sub TraceLine(ByVal txt As String)
    If Len(gTracePath) = 0 Then Exit Sub
    Dim fnum As Integer
    On Error GoTo Fail
    fnum = FreeFile
    Open gTracePath For Append As #fnum
    Print #fnum, Format$(Now, "hh:nn:ss") & " " & txt
    Close #fnum
    Exit Sub
Fail:
    ' Tracing must never break the workflow.
End Sub

' ---- Sheet plumbing ------------------------------------------------------

Public Function SheetByName(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    Set SheetByName = ws
End Function

Public Function SheetExists(ByVal sheetName As String) As Boolean
    SheetExists = Not SheetByName(sheetName) Is Nothing
End Function

' The shared Sites table - the single source of truth (F2).
Public Function SitesSheet() As Worksheet
    Dim ws As Worksheet
    Set ws = SheetByName(SH_SITES)
    If ws Is Nothing Then
        Err.Raise vbObjectError + 1, "SitesSheet", _
            "The '" & SH_SITES & "' sheet is missing. Click Build / Reset Workbook on the Start Here sheet first."
    End If
    Set SitesSheet = ws
End Function

' Last data row in Sites, judged by Site Name OR a coordinate being present.
Public Function SitesLastRow() As Long
    Dim ws As Worksheet, r As Long, lastUsed As Long
    Set ws = SitesSheet()
    lastUsed = ws.Cells(ws.Rows.Count, COL_SITENAME).End(xlUp).Row
    Dim lastCoord As Long
    lastCoord = ws.Cells(ws.Rows.Count, COL_LAT).End(xlUp).Row
    If lastCoord > lastUsed Then lastUsed = lastCoord
    If lastUsed < SITES_FIRST_DATA_ROW Then
        SitesLastRow = SITES_FIRST_DATA_ROW - 1   ' no data rows
    Else
        SitesLastRow = lastUsed
    End If
End Function

' True when a row has no site name and no coordinates - treated as empty.
Public Function RowIsEmpty(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    RowIsEmpty = IsBlank(ws.Cells(r, COL_SITENAME).Value) _
        And IsBlank(ws.Cells(r, COL_LAT).Value) _
        And IsBlank(ws.Cells(r, COL_LON).Value) _
        And IsBlank(ws.Cells(r, COL_ADDRESS).Value)
End Function

Public Function IsBlank(ByVal v As Variant) As Boolean
    IsBlank = (Len(Trim$(CStr(v))) = 0)
End Function

' Substitute {LAT}/{LON} into a URL template. Coordinates are written with
' invariant (US) decimal points so a comma decimal locale cannot corrupt them.
Public Function BuildUrl(ByVal template As String, ByVal lat As Variant, ByVal lon As Variant) As String
    Dim s As String
    s = Replace(template, "{LAT}", InvariantNum(lat))
    s = Replace(s, "{LON}", InvariantNum(lon))
    BuildUrl = s
End Function

Public Function InvariantNum(ByVal v As Variant) As String
    Dim s As String
    s = Trim$(CStr(v))
    InvariantNum = Replace(s, ",", ".")
End Function

' Coordinates valid? (lat -90..90, lon -180..180, both numeric and present)
Public Function HasValidCoords(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim la As Variant, lo As Variant
    la = ws.Cells(r, COL_LAT).Value
    lo = ws.Cells(r, COL_LON).Value
    If IsBlank(la) Or IsBlank(lo) Then Exit Function
    If Not IsNumeric(la) Or Not IsNumeric(lo) Then Exit Function
    If CDbl(la) < -90 Or CDbl(la) > 90 Then Exit Function
    If CDbl(lo) < -180 Or CDbl(lo) > 180 Then Exit Function
    HasValidCoords = True
End Function

Public Sub SetStatus(ByVal msg As String)
    Application.StatusBar = msg
End Sub

Public Sub ClearStatus()
    Application.StatusBar = False
End Sub

' Read a Setup named-range value as trimmed text ("" if missing/blank).
Public Function SetupValue(ByVal namedRange As String) As String
    Dim rng As Range
    On Error Resume Next
    Set rng = ThisWorkbook.Names(namedRange).RefersToRange
    On Error GoTo 0
    If rng Is Nothing Then
        SetupValue = ""
    Else
        SetupValue = Trim$(CStr(rng.Value))
    End If
End Function

' Convert a 1-based column index to its letter(s): 1->A, 27->AA.
Public Function ColLetter(ByVal n As Long) As String
    Dim r As String, m As Long
    Do While n > 0
        m = (n - 1) Mod 26
        r = Chr$(65 + m) & r
        n = (n - 1) \ 26
    Loop
    ColLetter = r
End Function

' Sleep n seconds without blocking other Excel work.
Public Sub WaitSeconds(ByVal seconds As Long)
    If seconds <= 0 Then Exit Sub
    Application.Wait Now + TimeSerial(0, 0, seconds)
End Sub

' Join a WO and DI into one display/filename string with a chosen
' separator and per-id prefix. Each piece is omitted when its value is
' blank, and the separator is only inserted between two present pieces:
'   JobIds("123", "456", " ", "WO", "DI")  -> "WO123 DI456"
'   JobIds("",    "456", " ", "WO", "DI")  -> "DI456"
'   JobIds("123", "",    "-", "WO", "DI")  -> "WO123"
'   JobIds("",    "",    " ", "WO", "DI")  -> ""
' Used by the FIRMette / Map filename builders, the default output folder
' path, and the WO #... map-page textbox so a missing WO or DI never
' leaves a dangling "WO " or trailing "-DI" in the output.
Public Function JobIds(ByVal wo As String, ByVal di As String, _
        ByVal sep As String, ByVal woPrefix As String, _
        ByVal diPrefix As String) As String
    Dim out As String
    wo = Trim$(wo): di = Trim$(di)
    If Len(wo) > 0 Then out = woPrefix & wo
    If Len(di) > 0 Then
        If Len(out) > 0 Then out = out & sep
        out = out & diPrefix & di
    End If
    JobIds = out
End Function

' Strip characters that are illegal in Windows file names.
Public Function CleanFileName(ByVal s As String) As String
    Dim bad As Variant, ch As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In bad
        s = Replace(s, CStr(ch), "_")
    Next ch
    CleanFileName = Trim$(s)
End Function
