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
Public Function ExtractStrings(ByVal json As String, ByVal fieldName As String) As Collection
    Dim col As New Collection, re As Object, m As Object
    Set re = NewRegex("""" & fieldName & """\s*:\s*""((?:[^""\\]|\\.)*)""", True)
    For Each m In re.Execute(json)
        col.Add JsonUnescape(m.SubMatches(0))
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

' GET a URL and write its body to disk as binary (PDF). Returns True on
' success. Verifies the response is actually a PDF by checking
' Content-Type — FEMA's GP service has been known to return an HTML error
' page with status 200, which would otherwise produce a corrupted .pdf.
Public Function HttpDownloadPdf(ByVal url As String, ByVal fullPath As String, _
        Optional ByRef errMsg As String) As Boolean
    Dim http As Object, stm As Object, contentType As String
    errMsg = ""
    TraceLine "HTTP GET (PDF) " & Left$(url, 200)
    On Error GoTo Fail
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 30000, 60000, 60000, 120000
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", BROWSER_UA
    http.setRequestHeader "Accept", "application/pdf,application/octet-stream,*/*"
    http.send
    If CLng(http.Status) <> 200 Then
        errMsg = "HTTP " & http.Status
        TraceLine "  -> " & errMsg
        Exit Function
    End If
    contentType = LCase$(CStr(http.getResponseHeader("Content-Type")))
    If InStr(1, contentType, "pdf", vbTextCompare) = 0 Then
        errMsg = "Response was not a PDF (Content-Type=" & contentType & ")"
        TraceLine "  -> " & errMsg
        Exit Function
    End If
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1                       ' adTypeBinary
    stm.Open
    stm.Write http.responseBody
    stm.SaveToFile fullPath, 2          ' adSaveCreateOverWrite
    stm.Close
    HttpDownloadPdf = True
    TraceLine "  -> 200 PDF saved (" & Len(http.responseBody) & " bytes)"
    Exit Function
Fail:
    errMsg = Err.Description
    TraceLine "  -> EXCEPTION: " & errMsg
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
