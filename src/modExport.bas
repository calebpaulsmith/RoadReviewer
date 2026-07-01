Attribute VB_Name = "modExport"
Option Explicit

' RoadReviewer V1 - CSV export of the Sites table including lookup results (F10).
' Writes every header + data row to a UTF-8 CSV in the output folder so a
' reviewer can take the whole hand-off in one file.

Public Sub ExportSitesCsv()
    Dim ws As Worksheet, last As Long, r As Long, c As Long, line As String, csv As String
    Set ws = SitesSheet()
    last = SitesLastRow()
    If last < SITES_FIRST_DATA_ROW Then
        If Not gHeadless Then MsgBox "No site rows to export.", vbInformation, "Export CSV"
        Exit Sub
    End If

    ' Header row.
    For c = 1 To COL_LAST
        line = line & IIf(c > 1, ",", "") & CsvField(CStr(ws.Cells(SITES_HEADER_ROW, c).Value))
    Next c
    csv = line & vbCrLf

    ' Data rows (skip fully-empty rows).
    For r = SITES_FIRST_DATA_ROW To last
        If Not RowIsEmpty(ws, r) Then
            line = ""
            For c = 1 To COL_LAST
                line = line & IIf(c > 1, ",", "") & CsvField(CellText(ws, r, c))
            Next c
            csv = csv & line & vbCrLf
        End If
    Next r

    Dim folder As String, file As String
    folder = ResolveOutputFolder()
    If Not EnsureFolderExists(folder) Then
        If Not gHeadless Then MsgBox "Could not create the output folder:" & vbCrLf & folder, vbExclamation, "Export CSV"
        Exit Sub
    End If
    file = folder & "RoadReviewer Sites.csv"

    If Not WriteCsvFile(file, csv) Then
        If Not gHeadless Then MsgBox "Could not write the CSV (is it open in another program?).", vbExclamation, "Export CSV"
        Exit Sub
    End If
    If Not gHeadless Then MsgBox "Exported the Sites table to:" & vbCrLf & file, vbInformation, "Export CSV"
End Sub

' For hyperlink-formula columns, export the resolved URL rather than "Map".
Private Function CellText(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    Select Case c
        Case COL_GMAP: CellText = BuildUrl(URL_GMAP, ws.Cells(r, COL_LAT).Value, ws.Cells(r, COL_LON).Value)
        Case COL_STREETVIEW: CellText = BuildUrl(URL_STREETVIEW, ws.Cells(r, COL_LAT).Value, ws.Cells(r, COL_LON).Value)
        Case COL_BING: CellText = BuildUrl(URL_BING, ws.Cells(r, COL_LAT).Value, ws.Cells(r, COL_LON).Value)
        Case COL_FEMAVIEW: CellText = BuildUrl(URL_FEMAVIEW, ws.Cells(r, COL_LAT).Value, ws.Cells(r, COL_LON).Value)
        Case COL_FIRMPORTAL: CellText = BuildUrl(URL_FIRMPORTAL, ws.Cells(r, COL_LAT).Value, ws.Cells(r, COL_LON).Value)
        Case COL_NFCMAP: CellText = NfcMapUrlForRow(ws, r)
        Case COL_AGOLMAP: CellText = AgolUrlForRow(ws, r)
        Case Else: CellText = CStr(ws.Cells(r, c).Value)
    End Select
    ' Blank the link columns when the row has no coordinates.
    Select Case c
        Case COL_GMAP, COL_STREETVIEW, COL_BING, COL_FEMAVIEW, COL_FIRMPORTAL, COL_NFCMAP, COL_AGOLMAP
            If Not HasValidCoords(ws, r) Then CellText = ""
    End Select
End Function

' Resolve the state-specific NFC map link (F8/F11). Matches the cell
' formula generated in SetNfcMapFormula - keep the two in sync.
Private Function NfcMapUrlForRow(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim stateCode As String, template As String
    stateCode = UCase$(SetupValue(NR_STATE))
    If Len(stateCode) = 0 Then stateCode = "MI"   ' matches ClassifyRows's default
    Select Case stateCode
        Case "MI": template = URL_NFC_MAPVIEW
        Case "IN": template = URL_NFC_MAPVIEW_IN
        Case "WI": template = URL_NFC_MAPVIEW_WI
        Case Else: template = URL_FEMAVIEW
    End Select
    NfcMapUrlForRow = BuildUrl(template, ws.Cells(r, COL_LAT).Value, ws.Cells(r, COL_LON).Value)
End Function

' Resolve the inspector's pasted AGOL webmap URL into a per-row deep-link
' that centers the map on this row's coords. Matches the cell formula
' generated in SetAgolMapFormula. Returns "" when NR_AGOLMAP is blank.
Private Function AgolUrlForRow(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim base As String, sep As String, lat As String, lon As String
    base = SetupValue(NR_AGOLMAP)
    If Len(base) = 0 Then Exit Function
    If Not HasValidCoords(ws, r) Then Exit Function
    sep = IIf(InStr(base, "?") > 0, "&", "?")
    lat = InvariantNum(ws.Cells(r, COL_LAT).Value)
    lon = InvariantNum(ws.Cells(r, COL_LON).Value)
    AgolUrlForRow = base & sep & "center=" & lon & "," & lat & _
        "&level=16&marker=" & lon & "," & lat
End Function

Private Function CsvField(ByVal s As String) As String
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Or InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Then
        CsvField = """" & Replace(s, """", """""") & """"
    Else
        CsvField = s
    End If
End Function

Private Function WriteCsvFile(ByVal path As String, ByVal content As String) As Boolean
    Dim fso As Object, ts As Object
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(path, True, False)   ' overwrite, ANSI (Excel-friendly CSV)
    ts.Write content
    ts.Close
    WriteCsvFile = True
    Exit Function
Fail:
    WriteCsvFile = False
End Function
