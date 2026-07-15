Attribute VB_Name = "modHttp"
Option Explicit

' RoadReviewer V1 - HTTP + minimal JSON helpers.
' Network stack per N1: MSXML2.ServerXMLHTTP.6.0 for GET, VBScript.RegExp for
' field extraction. JSON parsing is intentionally narrow - we only pull a
' handful of known scalar fields out of ArcGIS "attributes" blocks, which keeps
' it robust against the nesting/escaping issues that bite a general parser
' (friction point #3).

' GET a URL as text. Always sends a browser User-Agent because MDOT returns
' HTTP 403 to the default MSXML UA (§4.2 operational note); harmless elsewhere.
' Returns "" on any error and sets errMsg.
Public Function HttpGetText(ByVal url As String, Optional ByRef errMsg As String) As String
    Dim http As Object
    errMsg = ""
    TraceLine "HTTP GET " & Left$(url, 400)
    On Error GoTo Fail
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 5000, 20000, 30000
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", BROWSER_UA
    http.setRequestHeader "Accept", "application/json, text/plain, */*"
    http.send
    If http.Status = 200 Then
        HttpGetText = http.responseText
        TraceLine "  -> 200 (" & Len(HttpGetText) & " bytes)"
    Else
        errMsg = "HTTP " & http.Status
        TraceLine "  -> " & errMsg
    End If
    Exit Function
Fail:
    errMsg = Err.Description
    TraceLine "  -> EXCEPTION: " & errMsg
    HttpGetText = ""
End Function

' Number of features in an ArcGIS query response (counts "attributes" blocks).
Public Function FeatureCount(ByVal json As String) As Long
    FeatureCount = CountMatches(json, """attributes""\s*:")
End Function

' All integer values for a numeric field across every feature, e.g.
' "FunctionalSystem":6  ->  one entry per occurrence.
Public Function ExtractIntegers(ByVal json As String, ByVal fieldName As String) As Collection
    Dim col As New Collection, re As Object, m As Object
    Set re = NewRegex("""" & fieldName & """\s*:\s*(-?\d+)", True)
    For Each m In re.Execute(json)
        col.Add CLng(m.SubMatches(0))
    Next m
    Set ExtractIntegers = col
End Function

' All string values for a string field across every feature.
' Skips the self-referential `"FIELD":"FIELD"` mapping that some ArcGIS
' servers (TIGER) include in their `fieldAliases` block — that match
' would otherwise contaminate results with the literal field name.
Public Function ExtractStrings(ByVal json As String, ByVal fieldName As String) As Collection
    Dim col As New Collection, re As Object, m As Object, v As String
    Set re = NewRegex("""" & fieldName & """\s*:\s*""((?:[^""\\]|\\.)*)""", True)
    For Each m In re.Execute(json)
        v = JsonUnescape(m.SubMatches(0))
        If v <> fieldName Then col.Add v
    Next m
    Set ExtractStrings = col
End Function

' First string value for a field, "" if absent.
Public Function FirstString(ByVal json As String, ByVal fieldName As String) As String
    Dim c As Collection
    Set c = ExtractStrings(json, fieldName)
    If c.Count > 0 Then FirstString = c(1)
End Function

' True if the ArcGIS response carries an {"error":{...}} envelope.
Public Function HasArcgisError(ByVal json As String) As Boolean
    HasArcgisError = (InStr(json, """error""") > 0 And InStr(json, """code""") > 0)
End Function

' ---- Per-feature parsing + point-to-road distance (§ road distances) ------
' The scalar extractors above flatten a field across ALL features. To pair
' each road segment's class/name WITH its geometry (for distances) we split
' the response into one text block per feature. ArcGIS emits each feature as
' {"attributes":{...},"geometry":{...}} in that order, so splitting on the
' "attributes" key yields one chunk per feature that carries that feature's
' attributes AND the geometry that follows (up to the next feature).

Public Function FeatureBlocks(ByVal json As String) As Collection
    Dim parts() As String, i As Long, col As New Collection
    parts = Split(json, """attributes""")
    For i = 1 To UBound(parts)      ' parts(0) is the response preamble
        col.Add parts(i)
    Next i
    Set FeatureBlocks = col
End Function

' Minimum distance (in FEET) from the point (lonP,latP, decimal degrees) to
' the polyline geometry in one feature block. Geometry must be requested with
' outSR=4326 so vertices are lon/lat. Uses a local equirectangular projection
' (accurate at the few-hundred-foot scale we care about) and true
' point-to-segment distance, not just nearest vertex (a long straight segment
' can pass close to the point between far-apart vertices). Returns a huge
' number when the block has no geometry.
Public Function MinDistanceFt(ByVal block As String, ByVal lonP As Double, ByVal latP As Double) As Double
    Dim geoPos As Long, geom As String
    geoPos = InStr(block, """geometry""")
    If geoPos = 0 Then MinDistanceFt = 1E+15: Exit Function
    geom = Mid$(block, geoPos)

    Dim re As Object, matches As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "\[\s*(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)\s*\]"
    re.Global = True
    Set matches = re.Execute(geom)
    If matches.Count = 0 Then MinDistanceFt = 1E+15: Exit Function

    Const M_PER_DEG As Double = 111320#
    Const PI As Double = 3.14159265358979
    Dim mPerLon As Double, mPerLat As Double
    mPerLat = M_PER_DEG
    mPerLon = M_PER_DEG * Cos(latP * PI / 180#)

    ' The point sits at the local origin; project each vertex to metres
    ' relative to it, then take the min point-to-segment distance.
    Dim n As Long, i As Long
    n = matches.Count
    Dim px() As Double, py() As Double
    ReDim px(0 To n - 1): ReDim py(0 To n - 1)
    For i = 0 To n - 1
        Set m = matches(i)
        px(i) = (CDbl(m.SubMatches(0)) - lonP) * mPerLon
        py(i) = (CDbl(m.SubMatches(1)) - latP) * mPerLat
    Next i

    Dim best As Double, d As Double
    best = 1E+15
    If n = 1 Then
        best = Sqr(px(0) * px(0) + py(0) * py(0))
    Else
        For i = 0 To n - 2
            d = PointSegDistM(px(i), py(i), px(i + 1), py(i + 1))
            If d < best Then best = d
        Next i
    End If
    MinDistanceFt = best * 3.28084          ' metres -> feet
End Function

' Distance from the origin (0,0) to the segment A(ax,ay)-B(bx,by), all in
' the same planar units (metres). Standard proj(point onto segment) clamp.
Private Function PointSegDistM(ByVal ax As Double, ByVal ay As Double, _
        ByVal bx As Double, ByVal by As Double) As Double
    Dim dx As Double, dy As Double, l2 As Double, t As Double, cx As Double, cy As Double
    dx = bx - ax: dy = by - ay
    l2 = dx * dx + dy * dy
    If l2 = 0# Then
        PointSegDistM = Sqr(ax * ax + ay * ay)
        Exit Function
    End If
    t = -(ax * dx + ay * dy) / l2          ' projection of origin onto AB
    If t < 0# Then t = 0#                  ' single-line If: no ElseIf allowed
    If t > 1# Then t = 1#
    cx = ax + t * dx: cy = ay + t * dy
    PointSegDistM = Sqr(cx * cx + cy * cy)
End Function

' `global` is a VBA reserved word (synonym for Public), so use `isGlobal`
' for the parameter — using the reserved word as a parameter name has bitten
' us with "Sub or Function not defined" errors at JIT-compile time even
' though the .bas file imports cleanly.
' IgnoreCase defaults to False because ArcGIS responses include both a
' lower-case `"name"` key inside the fields-metadata block (where the value
' is the field's name like "OBJECTID") AND an upper-case attribute key
' (e.g. `"NAME":"Kalamazoo, MI"`). If our regex were case-insensitive,
' FirstString("NAME") would return "OBJECTID" — the first lowercase
' `"name":...` match — instead of the urban-area NAME we actually want.
Private Function NewRegex(ByVal pattern As String, ByVal isGlobal As Boolean, _
        Optional ByVal ignoreCase As Boolean = False) As Object
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = ignoreCase
    re.global = isGlobal
    Set NewRegex = re
End Function

Private Function CountMatches(ByVal s As String, ByVal pattern As String) As Long
    Dim re As Object
    Set re = NewRegex(pattern, True)
    CountMatches = re.Execute(s).Count
End Function

' GET a URL and write its body to disk as binary. Returns True on success.
' Verifies the response Content-Type contains expectType (case-insensitive,
' skipped when "") — the FEMA GP service has been known to return an HTML
' error page with status 200, which would otherwise produce a corrupted
' file, and the Esri imagery export returns a JSON error the same way.
Public Function HttpDownloadBinary(ByVal url As String, ByVal fullPath As String, _
        ByVal acceptHeader As String, ByVal expectType As String, _
        Optional ByRef errMsg As String) As Boolean
    Dim http As Object, stm As Object, contentType As String
    errMsg = ""
    TraceLine "HTTP GET (binary) " & Left$(url, 200)
    On Error GoTo Fail
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 30000, 60000, 60000, 120000
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", BROWSER_UA
    http.setRequestHeader "Accept", acceptHeader
    http.send
    If CLng(http.Status) <> 200 Then
        errMsg = "HTTP " & http.Status
        TraceLine "  -> " & errMsg
        Exit Function
    End If
    contentType = LCase$(CStr(http.getResponseHeader("Content-Type")))
    If Len(expectType) > 0 Then
        If InStr(1, contentType, expectType, vbTextCompare) = 0 Then
            errMsg = "Response was not " & expectType & " (Content-Type=" & contentType & ")"
            TraceLine "  -> " & errMsg
            Exit Function
        End If
    End If
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                       ' adTypeBinary
    stm.Open
    stm.Write http.responseBody
    stm.SaveToFile fullPath, 2          ' adSaveCreateOverWrite
    stm.Close
    HttpDownloadBinary = True
    TraceLine "  -> 200 saved (" & Len(http.responseBody) & " bytes)"
    Exit Function
Fail:
    errMsg = Err.Description
    TraceLine "  -> EXCEPTION: " & errMsg
End Function

' Back-compat wrapper for the FIRMette flow.
Public Function HttpDownloadPdf(ByVal url As String, ByVal fullPath As String, _
        Optional ByRef errMsg As String) As Boolean
    HttpDownloadPdf = HttpDownloadBinary(url, fullPath, _
        "application/pdf,application/octet-stream,*/*", "pdf", errMsg)
End Function

Public Function JsonUnescape(ByVal s As String) As String
    s = Replace(s, "\""", """")
    s = Replace(s, "\/", "/")
    s = Replace(s, "\\", "\")
    s = Replace(s, "\n", vbLf)
    s = Replace(s, "\t", vbTab)
    s = Replace(s, "\r", vbCr)
    JsonUnescape = s
End Function

' URL-encode a value for a query string (the bits ArcGIS/Census care about).
Public Function UrlEncode(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) _
            Or (code >= 97 And code <= 122) Or InStr("-_.~", ch) > 0 Then
            out = out & ch
        ElseIf ch = " " Then
            out = out & "+"
        Else
            out = out & "%" & Right$("0" & Hex$(code), 2)
        End If
    Next i
    UrlEncode = out
End Function
