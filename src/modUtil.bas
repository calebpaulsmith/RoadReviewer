Attribute VB_Name = "modUtil"
Option Explicit

' RoadReviewer V1 - shared helpers used across workflows.
' Every sheet reference is hard-bound by name (friction point #4): no
' workflow ever writes to ActiveSheet.

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
            "The '" & SH_SITES & "' sheet is missing. Run BuildWorkbook from the Home sheet first."
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

' Strip characters that are illegal in Windows file names.
Public Function CleanFileName(ByVal s As String) As String
    Dim bad As Variant, ch As Variant
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In bad
        s = Replace(s, CStr(ch), "_")
    Next ch
    CleanFileName = Trim$(s)
End Function
