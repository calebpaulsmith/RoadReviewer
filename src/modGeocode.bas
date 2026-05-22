Attribute VB_Name = "modGeocode"
Option Explicit

' RoadReviewer V1 - address geocoding (F4).
' For rows that have an Address but no coordinates, calls the US Census Bureau
' one-line geocoder (free, no auth) and fills Latitude/Longitude. Coordinates
' already present are NEVER overwritten.

Public Sub GeocodeAddresses()
    Dim ws As Worksheet, last As Long, r As Long
    Dim total As Long, done As Long, ok As Long

    Set ws = SitesSheet()
    last = SitesLastRow()
    If last < SITES_FIRST_DATA_ROW Then
        MsgBox "No site rows found. Add addresses on the Sites sheet first.", vbInformation, "Geocode"
        Exit Sub
    End If

    For r = SITES_FIRST_DATA_ROW To last
        If NeedsGeocode(ws, r) Then total = total + 1
    Next r
    If total = 0 Then
        MsgBox "No rows need geocoding (every row either has coordinates already or has no address).", _
            vbInformation, "Geocode"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    For r = SITES_FIRST_DATA_ROW To last
        If NeedsGeocode(ws, r) Then
            done = done + 1
            SetStatus "Geocoding " & done & " of " & total & " - " & CStr(ws.Cells(r, COL_ADDRESS).Value)
            If GeocodeOneRow(ws, r) Then ok = ok + 1
            DoEvents
        End If
    Next r
    Application.ScreenUpdating = True
    ClearStatus
    MsgBox "Geocoded " & ok & " of " & total & " address(es).", vbInformation, "Geocode"
End Sub

' A row needs geocoding when it has an address but no valid coordinates.
Private Function NeedsGeocode(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    If RowIsEmpty(ws, r) Then Exit Function
    If HasValidCoords(ws, r) Then Exit Function
    NeedsGeocode = Not IsBlank(ws.Cells(r, COL_ADDRESS).Value)
End Function

Private Function GeocodeOneRow(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim addr As String, url As String, json As String, errMsg As String
    addr = Trim$(CStr(ws.Cells(r, COL_ADDRESS).Value))
    url = REST_CENSUS_GEOCODE & "?address=" & UrlEncode(addr) & _
        "&benchmark=Public_AR_Current&format=json"
    json = HttpGetText(url, errMsg)
    If Len(errMsg) > 0 Then
        ws.Cells(r, COL_GEOCODE).Value = STATUS_FAILED_PREFIX & errMsg
        Exit Function
    End If

    ' Census returns coordinates.x = longitude, coordinates.y = latitude.
    Dim lon As String, lat As String
    lon = FirstNumberAfter(json, """x""")
    lat = FirstNumberAfter(json, """y""")
    If Len(lat) = 0 Or Len(lon) = 0 Then
        ws.Cells(r, COL_GEOCODE).Value = STATUS_FAILED_PREFIX & "no match"
        Exit Function
    End If

    ws.Cells(r, COL_LAT).Value = CDbl(lat)
    ws.Cells(r, COL_LON).Value = CDbl(lon)
    ws.Cells(r, COL_GEOCODE).Value = "Geocoded"
    GeocodeOneRow = True
End Function

' Pull the first JSON number that follows a key like "x" or "y".
Private Function FirstNumberAfter(ByVal json As String, ByVal key As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = key & "\s*:\s*(-?\d+(?:\.\d+)?)"
    re.IgnoreCase = True
    re.global = False
    If re.Test(json) Then
        Set m = re.Execute(json)(0)
        FirstNumberAfter = m.SubMatches(0)
    End If
End Function
