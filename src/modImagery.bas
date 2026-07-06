Attribute VB_Name = "modImagery"
Option Explicit

' RoadReviewer V1 - Workflow 2: Review Imagery (F3.2).
' Opens a curated set of imagery URLs in the default browser for each selected
' Sites row. Select the row(s) on the Sites sheet, then run.

Private Const MAX_AUTO_OPEN As Long = 8   ' confirm above this many rows

Public Sub OpenImageryForSelection()
    Dim ws As Worksheet
    Set ws = SitesSheet()
    ws.Activate   ' restores the user's last selection on Sites

    If Not (TypeOf Selection Is Range) Then
        MsgBox "Select one or more rows in the Sites table (click the row numbers), then run this.", _
            vbInformation, "Review Imagery"
        Exit Sub
    End If

    Dim rows As Collection
    Set rows = SelectedDataRows(ws, Selection)
    If rows.Count = 0 Then
        MsgBox "No rows with coordinates are selected. Click a site row (or run Classify/geocode first), then try again.", _
            vbInformation, "Review Imagery"
        Exit Sub
    End If

    If rows.Count > MAX_AUTO_OPEN Then
        If MsgBox("This will open about " & rows.Count * 5 & " browser tabs (" & rows.Count & _
            " sites x 5 sources). Continue?", vbQuestion + vbYesNo, "Review Imagery") <> vbYes Then Exit Sub
    End If

    Dim r As Variant, opened As Long
    For Each r In rows
        OpenImageryForRow ws, CLng(r)
        opened = opened + 1
        SetStatus "Opening imagery " & opened & " of " & rows.Count & "..."
        DoEvents
    Next r
    ClearStatus
End Sub

Private Function SelectedDataRows(ByVal ws As Worksheet, ByVal sel As Range) As Collection
    Dim col As New Collection, area As Range, r As Long
    Dim seen As String
    For Each area In sel.Areas
        For r = area.Row To area.Row + area.rows.Count - 1
            If r >= SITES_FIRST_DATA_ROW Then
                If InStr(seen, "|" & r & "|") = 0 Then
                    If HasValidCoords(ws, r) Then
                        col.Add r
                        seen = seen & "|" & r & "|"
                    End If
                End If
            End If
        Next r
    Next area
    Set SelectedDataRows = col
End Function

Private Sub OpenImageryForRow(ByVal ws As Worksheet, ByVal r As Long)
    Dim lat As Variant, lon As Variant
    lat = ws.Cells(r, COL_LAT).Value
    lon = ws.Cells(r, COL_LON).Value
    OpenUrl BuildUrl(URL_GMAP, lat, lon)
    OpenUrl BuildUrl(URL_STREETVIEW, lat, lon)
    OpenUrl BuildUrl(URL_BING, lat, lon)
    OpenUrl BuildUrl(URL_GEARTH, lat, lon)
    OpenUrl BuildUrl(URL_FEMAVIEW, lat, lon)
End Sub

Private Sub OpenUrl(ByVal url As String)
    On Error Resume Next
    ThisWorkbook.FollowHyperlink Address:=url, NewWindow:=False
    On Error GoTo 0
End Sub
