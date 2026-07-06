Attribute VB_Name = "modGeocode"
Option Explicit

' Address geocoding (F4). No standalone button anymore - Check Roads
' (modClassify.ClassifyOneRow) calls GeocodeRow for any row that has an
' Address but no coordinates, so there is no geocode-then-classify ordering
' for the user to get wrong. Coordinates already present are NEVER
' overwritten (the caller only invokes this when HasValidCoords is False,
' and the function re-checks).

' Geocode one row's Address into Latitude/Longitude via the US Census
' Bureau one-line geocoder (free, no auth). Returns True on success;
' False with errMsg set (and a Failed marker in the Geocode Status
' column) otherwise.
Public Function GeocodeRow(ByVal ws As Worksheet, ByVal r As Long, _
        ByRef errMsg As String) As Boolean
    If HasValidCoords(ws, r) Then
        GeocodeRow = True   ' nothing to do; never overwrite typed coords
        Exit Function
    End If

    Dim addr As String, url As String, json As String, httpErr As String
    addr = Trim$(CStr(ws.Cells(r, COL_ADDRESS).Value))
    If Len(addr) = 0 Then
        errMsg = "no address"
        ws.Cells(r, COL_GEOCODE).Value = STATUS_FAILED_PREFIX & errMsg
        Exit Function
    End If

    TraceLine "GeocodeRow row=" & r & " addr=" & addr
    url = REST_CENSUS_GEOCODE & "?address=" & UrlEncode(addr) & _
        "&benchmark=Public_AR_Current&format=json"
    json = HttpGetText(url, httpErr)
    If Len(httpErr) > 0 Then
        errMsg = httpErr
        ws.Cells(r, COL_GEOCODE).Value = STATUS_FAILED_PREFIX & errMsg
        Exit Function
    End If

    ' Census returns coordinates.x = longitude, coordinates.y = latitude.
    Dim lon As String, lat As String
    lon = FirstNumberAfter(json, """x""")
    lat = FirstNumberAfter(json, """y""")
    If Len(lat) = 0 Or Len(lon) = 0 Then
        errMsg = "no match for address"
        ws.Cells(r, COL_GEOCODE).Value = STATUS_FAILED_PREFIX & errMsg
        Exit Function
    End If

    ws.Cells(r, COL_LAT).Value = CDbl(lat)
    ws.Cells(r, COL_LON).Value = CDbl(lon)
    ws.Cells(r, COL_GEOCODE).Value = "Geocoded"
    GeocodeRow = True
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
