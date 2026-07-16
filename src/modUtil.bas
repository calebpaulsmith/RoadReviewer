Attribute VB_Name = "modUtil"
Option Explicit

' RoadReviewer V1 - shared helpers used across workflows.
' Every sheet reference is hard-bound by name (friction point #4): no
' workflow ever writes to ActiveSheet.

' ---- Module-level state ---------------------------------------------------
' All Public variables must live in the declarations section (before the
' first Sub/Function) or VBA throws a compile error.

' Win32 for SurfaceFolder: raise an already-open Explorer window instead of
' spawning a duplicate. Declares belong in the declarations section too.
#If VBA7 Then
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
#Else
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
#End If
Private Const SW_RESTORE As Long = 9

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

' Name of the "hub" sheet for the current product: the standard product's
' visible "Start Here" landing, or the inspector's hidden "Tools and Exports"
' utility sheet. Every SheetByName/Worksheets() lookup of the hub goes through
' this so the two products can name the same-role sheet differently.
Public Function StartSheetName() As String
    If ProductIsInspector() Then
        StartSheetName = SH_TOOLS
    Else
        StartSheetName = SH_START
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

' Normalize a State dropdown value to its bare 2-letter code. The dropdown
' annotates unwired states ("MN (not wired)"), so take the text before the
' first space and upper-case it: "MI" -> "MI", "MN (not wired)" -> "MN".
' NOT named StateCode: callers use a local variable `stateCode`, and VBA
' identifiers are case-insensitive, so a same-named function is shadowed by
' the local and `stateCode = StateCode(...)` compiles as array-indexing the
' local ("Expected array") - the exact NfcWired/nfcWired trap noted in
' modClassify. Hence BareStateCode.
Public Function BareStateCode(ByVal raw As String) As String
    Dim s As String, p As Long
    s = Trim$(raw)
    p = InStr(s, " ")
    If p > 0 Then s = Left$(s, p - 1)
    BareStateCode = UCase$(Trim$(s))
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

' The disaster tag used in file names, pairing the Disaster Number with the
' State. Convention (user direction, PR #37): NO separators - "DR4882IN",
' not "DR-4882-IN". A bare number is assumed to be a major-disaster DR
' ("4882" -> "DR4882"); a typed prefix (DR/EM/...) is kept, with any hyphens
' or spaces the user typed stripped out. State is appended when set (unless
' the user already typed it). Empty when there's no disaster.
Public Function DisasterTag() As String
    Dim d As String, st As String
    d = Trim$(SetupValue(NR_DISASTER))
    st = BareStateCode(SetupValue(NR_STATE))
    If Len(d) = 0 Then Exit Function
    d = UCase$(Replace(Replace(d, "-", ""), " ", ""))
    If IsAllDigits(d) Then d = "DR" & d
    DisasterTag = d
    If Len(st) > 0 Then
        If Right$(DisasterTag, Len(st)) <> st Then DisasterTag = DisasterTag & st
    End If
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

' The SHARED file-name stem for every export (Location Map PDF, FIRMettes, KML,
' CSV, GeoJSON, ...), so they're named consistently and never collide:
'   "WO123 DI5 - DR-4882-IN"   (whatever job info is set)
' When no WO/DI/Disaster is set, falls back to a date-time stamp so repeated
' exports don't overwrite one another. Each export appends its own suffix, e.g.
'   JobFileStem() & " - Location Map.pdf"
'   JobFileStem() & " - Sites.kml"
Public Function JobFileStem() As String
    JobFileStem = JobFileStemFor(SetupValue(NR_WO), SetupValue(NR_DI))
End Function

' Same stem, but with explicit WO/DI - so a per-site export (FIRMette, per-site
' Location Map) can pass a row's OWN WO/DI, which override the Setup values for
' that row (F14 / §9.5, same rule the stamp already applies). The Disaster tag
' stays job-wide (there is no per-row Disaster cell). Blank wo/di collapse away
' via JobIds exactly as the Setup-level path does.
Public Function JobFileStemFor(ByVal wo As String, ByVal di As String) As String
    Dim jobs As String, tag As String, stem As String
    jobs = JobIds(wo, di, " ", "WO", "DI")
    tag = DisasterTag()
    stem = jobs
    If Len(tag) > 0 Then
        If Len(stem) > 0 Then stem = stem & " - "
        stem = stem & tag
    End If
    If Len(stem) = 0 Then stem = Format$(Now(), "yyyy-mm-dd HHmm")
    JobFileStemFor = CleanFileName(stem)
End Function

' Show the output folder to the user after an export (user direction, PR #37:
' every output ends with its folder visible). If an Explorer window is already
' open AT that folder it is un-minimized and brought to the front; otherwise a
' new window opens. No-op when headless or when the folder doesn't exist.
Public Sub SurfaceFolder(ByVal folderPath As String)
    If gHeadless Then Exit Sub
    Dim target As String
    target = folderPath
    Do While Right$(target, 1) = "\"
        target = Left$(target, Len(target) - 1)
    Loop
    If Len(Dir$(target, vbDirectory)) = 0 Then Exit Sub

    ' Shell.Application.Windows enumerates open Explorer (and IE) windows;
    ' non-folder windows lack Document.Folder, hence the blanket error guard.
    On Error Resume Next
    Dim sh As Object, w As Object, p As String
    Set sh = CreateObject("Shell.Application")
    For Each w In sh.Windows
        p = ""
        p = w.Document.Folder.Self.Path
        If StrComp(p, target, vbTextCompare) = 0 Then
            ShowWindow w.hWnd, SW_RESTORE
            SetForegroundWindow w.hWnd
            Exit Sub
        End If
    Next w
    On Error GoTo 0

    Shell "explorer.exe """ & target & """", vbNormalFocus
End Sub

' Strip characters that are illegal in Windows file names.
Public Function CleanFileName(ByVal s As String) As String
    Dim bad As Variant, ch As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In bad
        s = Replace(s, CStr(ch), "_")
    Next ch
    CleanFileName = Trim$(s)
End Function
