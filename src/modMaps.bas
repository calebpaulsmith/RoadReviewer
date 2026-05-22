Attribute VB_Name = "modMaps"
Option Explicit

' RoadReviewer V1 - Workflow 3: Maps & FIRMettes (F3.3, F10).
' Output-folder resolution and KML export are implemented here. FIRMette
' download and Map Pages are ported in the next build increment - their
' buttons are wired to honest stubs so the workbook has no orphan controls.

' ---- output folder (§8.9) -------------------------------------------------

Public Sub SelectOutputFolder()
    Dim fd As Object, chosen As String
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "Choose the output folder for FIRMettes and map PDFs"
    If fd.Show = -1 Then
        chosen = fd.SelectedItems(1)
        If Right$(chosen, 1) <> "\" Then chosen = chosen & "\"
        ThisWorkbook.Names(NR_OUTFOLDER).RefersToRange.Value = chosen
    End If
End Sub

' Effective output folder: the Setup value if set, else the default pattern.
Public Function ResolveOutputFolder() As String
    Dim v As String
    v = SetupValue(NR_OUTFOLDER)
    If Len(v) > 0 Then
        If Right$(v, 1) <> "\" Then v = v & "\"
        ResolveOutputFolder = v
    Else
        ResolveOutputFolder = DefaultOutputFolder()
    End If
End Function

' {base}\Desktop\Script\RoadReviewer\{Disaster}\WO{WO}-DI{DI}\  (§8.9)
Private Function DefaultOutputFolder() As String
    Dim profile As String, base As String, disaster As String, wo As String, di As String
    profile = Environ$("USERPROFILE")
    If FolderExists(profile & "\OneDrive - FEMA") Then
        base = profile & "\OneDrive - FEMA"
    ElseIf FolderExists(profile & "\OneDrive") Then
        base = profile & "\OneDrive"
    Else
        base = profile
    End If
    disaster = CleanFileName(SetupValue(NR_DISASTER))
    wo = CleanFileName(SetupValue(NR_WO))
    di = CleanFileName(SetupValue(NR_DI))
    DefaultOutputFolder = base & "\Desktop\Script\RoadReviewer\" & _
        IIf(Len(disaster) > 0, disaster & "\", "") & _
        "WO" & wo & "-DI" & di & "\"
End Function

Private Function FolderExists(ByVal path As String) As Boolean
    FolderExists = (Len(Dir$(path, vbDirectory)) > 0)
End Function

' Create the folder (and any missing parents) on demand.
Public Function EnsureFolderExists(ByVal path As String) As Boolean
    Dim fso As Object
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    CreateTree fso, path
    EnsureFolderExists = fso.FolderExists(path)
    Exit Function
Fail:
    EnsureFolderExists = False
End Function

Private Sub CreateTree(ByVal fso As Object, ByVal path As String)
    Dim parent As String
    path = TrimSlash(path)
    If Len(path) = 0 Then Exit Sub
    If fso.FolderExists(path) Then Exit Sub
    parent = fso.GetParentFolderName(path)
    If Len(parent) > 0 And Not fso.FolderExists(parent) Then CreateTree fso, parent
    If Not fso.FolderExists(path) Then fso.CreateFolder path
End Sub

Private Function TrimSlash(ByVal s As String) As String
    Do While Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    TrimSlash = s
End Function

' ---- KML export (F10) -----------------------------------------------------

Public Sub ExportSitesToKML()
    Dim ws As Worksheet, last As Long, r As Long, kml As String, n As Long
    Set ws = SitesSheet()
    last = SitesLastRow()
    If last < SITES_FIRST_DATA_ROW Then
        MsgBox "No site rows to export.", vbInformation, "Export KML"
        Exit Sub
    End If

    kml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
        "<kml xmlns=""http://www.opengis.net/kml/2.2""><Document>" & vbCrLf & _
        "<name>RoadReviewer Sites</name>" & vbCrLf
    For r = SITES_FIRST_DATA_ROW To last
        If HasValidCoords(ws, r) Then
            kml = kml & PlacemarkXml(ws, r)
            n = n + 1
        End If
    Next r
    kml = kml & "</Document></kml>"

    If n = 0 Then
        MsgBox "No rows have valid coordinates to export.", vbInformation, "Export KML"
        Exit Sub
    End If

    Dim folder As String, file As String
    folder = ResolveOutputFolder()
    If Not EnsureFolderExists(folder) Then
        MsgBox "Could not create the output folder:" & vbCrLf & folder, vbExclamation, "Export KML"
        Exit Sub
    End If
    file = folder & "RoadReviewer Sites.kml"

    If Not WriteTextFile(file, kml) Then
        MsgBox "Could not write the KML file.", vbExclamation, "Export KML"
        Exit Sub
    End If

    Dim q As String: q = Chr$(34)
    Shell "cmd /c start " & q & q & " " & q & file & q, vbNormalFocus
    MsgBox "Exported " & n & " point(s) to:" & vbCrLf & file, vbInformation, "Export KML"
End Sub

Private Function PlacemarkXml(ByVal ws As Worksheet, ByVal r As Long) As String
    Dim nm As String, desc As String, lat As String, lon As String
    nm = XmlEscape(CStr(ws.Cells(r, COL_SITENAME).Value))
    If Len(nm) = 0 Then nm = "Site row " & r
    desc = XmlEscape(CStr(ws.Cells(r, COL_DESC).Value) & _
        IIf(IsBlank(ws.Cells(r, COL_ELIGIBILITY).Value), "", " | " & CStr(ws.Cells(r, COL_ELIGIBILITY).Value)))
    lat = InvariantNum(ws.Cells(r, COL_LAT).Value)
    lon = InvariantNum(ws.Cells(r, COL_LON).Value)
    PlacemarkXml = "<Placemark><name>" & nm & "</name>" & _
        IIf(Len(desc) > 0, "<description>" & desc & "</description>", "") & _
        "<Point><coordinates>" & lon & "," & lat & ",0</coordinates></Point></Placemark>" & vbCrLf
End Function

Private Function XmlEscape(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    XmlEscape = s
End Function

Private Function WriteTextFile(ByVal path As String, ByVal content As String) As Boolean
    Dim fso As Object, ts As Object
    On Error GoTo Fail
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(path, True, False)   ' overwrite, ANSI (ASCII-clean KML)
    ts.Write content
    ts.Close
    WriteTextFile = True
    Exit Function
Fail:
    WriteTextFile = False
End Function

' ---- not yet implemented (next build increment) ---------------------------

Public Sub DownloadFirmettes()
    NotYet "Download FIRMettes"
End Sub

Public Sub ReRunFailedFirmettes()
    NotYet "Re-run Failed FIRMettes"
End Sub

Public Sub PrepareMapPages()
    NotYet "Prepare Map Pages"
End Sub

Public Sub ExportCombinedMapPdf()
    NotYet "Export Combined Map PDF"
End Sub

Private Sub NotYet(ByVal feature As String)
    MsgBox feature & " is being ported from the prototype in the next build increment." & vbCrLf & _
        "Classify Roads, Review Imagery, geocoding, KML and CSV export are ready to use now.", _
        vbInformation, "RoadReviewer"
End Sub
