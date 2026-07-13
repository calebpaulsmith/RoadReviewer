Attribute VB_Name = "modExportMenu"
Option Explicit

' RoadReviewer - the Start Here "Export" dropdown.
'
' WHY A COMBO BOX + GO BUTTON, NOT A DATA-VALIDATION DROPDOWN
' -----------------------------------------------------------
' A Data Validation dropdown cannot run a macro when its value changes; that
' needs a Worksheet_Change handler, which lives in the sheet's CLASS module.
' This project imports standard .bas modules only, and docs/build-and-import.md
' specifically promises the build needs no "trust access to the VBA project
' object model" setting. Writing into a sheet class module would break that
' promise.
'
' A Form Control combo box has no such problem: modBuild drops it on Start
' Here, we read its selection with .ControlFormat.ListIndex, and an adjacent
' "Go" button calls RunSelectedExport. Zero event code, zero trust settings.
'
' Adding an export = one row in ExportItems(). The combo list and the
' dispatcher both read from it, so they can never drift apart.

Public Const EXPORT_COMBO_NAME As String = "RR_ExportPicker"
Public Const ROADS_COMBO_NAME As String = "RR_RoadsPicker"

' Caption <-> macro pairs, in menu order. Keep captions short; the combo is
' ~260pt wide. "|" separates caption from macro name.
'
' Product-aware: on the inspector product, KML, FIRMettes and the combined
' map PDF are dedicated workflow buttons on Start Here, so the dropdown only
' carries the general hand-off exports that have no button of their own. The
' standard product (no map/FIRMette workflow) keeps the full list.
Private Function ExportItems() As Variant
    If ProductIsInspector() Then
        ExportItems = Array( _
            "Sites Table (CSV)|ExportSitesCsv", _
            "Sites to GeoJSON|ExportSitesToGeoJson", _
            "Send Sites to AGOL Map (KML + open webmap)|SendSitesToAgolMap", _
            "Open Sites on NFC Layer (AGOL)|OpenSitesOnNfcLayer")
    Else
        ExportItems = Array( _
            "Sites Table (CSV)|ExportSitesCsv", _
            "Sites to KML|ExportSitesToKML", _
            "Sites to GeoJSON|ExportSitesToGeoJson", _
            "Combined Map PDF|ExportCombinedMapPdf", _
            "Download FIRMettes|DownloadFirmettes", _
            "Re-run Failed FIRMettes|ReRunFailedFirmettes", _
            "Send Sites to AGOL Map (KML + open webmap)|SendSitesToAgolMap", _
            "Open Sites on NFC Layer (AGOL)|OpenSitesOnNfcLayer")
    End If
End Function

' The classify + photo actions, collapsed off the Start Here surface into one
' dropdown (inspector) so the map/FIRMette workflows are the visible hero.
Private Function RoadsItems() As Variant
    RoadsItems = Array( _
        "Check roads (classify sites)|CheckRoads", _
        "Re-run failed rows|ReRunFailedRows", _
        "Open photo links (selected rows)|OpenImageryForSelection")
End Function

' Compile-check hook (build\compile-check.ps1): pure, no side effects. Calling
' it forces VBA to JIT-compile this whole module (including RoadsItems and the
' pickers below), surfacing any syntax error at build time instead of at the
' user's first "Go" click.
Public Function ExportItemCount() As Long
    Dim items As Variant
    items = ExportItems()
    ExportItemCount = UBound(items) - LBound(items) + 1
End Function

Private Function ItemCaption(ByVal item As String) As String
    ItemCaption = Left$(item, InStr(item, "|") - 1)
End Function

Private Function ItemMacro(ByVal item As String) As String
    ItemMacro = Mid$(item, InStr(item, "|") + 1)
End Function

' ---- built by modBuild ----------------------------------------------------

' Drops the export combo on ws and fills it from ExportItems().
Public Sub CreateExportPicker(ByVal ws As Worksheet, ByVal leftPt As Single, _
        ByVal topPt As Single, ByVal widthPt As Single)
    CreatePicker ws, EXPORT_COMBO_NAME, leftPt, topPt, widthPt, ExportItems()
End Sub

' Drops the roads & photos combo on ws and fills it from RoadsItems().
Public Sub CreateRoadsPicker(ByVal ws As Worksheet, ByVal leftPt As Single, _
        ByVal topPt As Single, ByVal widthPt As Single)
    CreatePicker ws, ROADS_COMBO_NAME, leftPt, topPt, widthPt, RoadsItems()
End Sub

' Generic: drop a named combo on ws and fill it. Idempotent - an existing
' picker of the same name is deleted first, so Build / Reset is safe.
Private Sub CreatePicker(ByVal ws As Worksheet, ByVal comboName As String, ByVal leftPt As Single, _
        ByVal topPt As Single, ByVal widthPt As Single, ByVal items As Variant)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name = comboName Then ws.Shapes(i).Delete
    Next i

    Dim shp As Shape, v As Variant
    Set shp = ws.Shapes.AddFormControl(xlDropDown, leftPt, topPt, widthPt, 22)
    shp.Name = comboName
    shp.Placement = xlFreeFloating

    With shp.ControlFormat
        .RemoveAllItems
        For Each v In items
            .AddItem ItemCaption(CStr(v))
        Next v
        .DropDownLines = UBound(items) - LBound(items) + 1
        .ListIndex = 1                      ' default to the first item
    End With
End Sub

' ---- wired to the "Go" buttons --------------------------------------------

Public Sub RunSelectedExport()
    RunPicker EXPORT_COMBO_NAME, ExportItems()
End Sub

Public Sub RunSelectedRoadsAction()
    RunPicker ROADS_COMBO_NAME, RoadsItems()
End Sub

' Reads the named picker on Start Here and runs the matching macro.
' Application.Run keeps this module free of hard references to modMaps /
' modExport / modClassify, so import order never matters.
Private Sub RunPicker(ByVal comboName As String, ByVal items As Variant)
    Dim ws As Worksheet, shp As Shape
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_START)
    Set shp = ws.Shapes(comboName)
    On Error GoTo 0

    If shp Is Nothing Then
        If Not gHeadless Then MsgBox _
            "The picker is missing. Click 'Build / Reset Workbook' to restore it.", _
            vbExclamation, "Action"
        Exit Sub
    End If

    Dim idx As Long
    idx = shp.ControlFormat.ListIndex
    If idx < 1 Then
        If Not gHeadless Then MsgBox "Choose an item from the dropdown first.", _
            vbInformation, "Action"
        Exit Sub
    End If

    If idx > (UBound(items) - LBound(items) + 1) Then Exit Sub
    Application.Run ItemMacro(CStr(items(LBound(items) + idx - 1)))
End Sub
