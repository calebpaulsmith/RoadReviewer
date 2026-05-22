Attribute VB_Name = "BuildHelper"
Option Explicit

' Build-time helper. Not part of the shipping workbook - the build script
' injects this module, runs BuildWorkbookSafe, then removes it before save.
' Captures any error from BuildWorkbook to a known temp file so the
' automation host can read the actual failure instead of staring at Excel
' sitting in VBE break mode.

Public Sub BuildWorkbookSafe()
    Dim logPath As String
    logPath = Environ$("TEMP") & "\RoadReviewer_build_error.txt"
    On Error GoTo Trap
    ' Wipe any prior error file so a stale one never lies to us.
    On Error Resume Next
    Kill logPath
    On Error GoTo Trap

    ' Tell BuildWorkbook to suppress its success/failure MsgBox.
    gSilentBuild = True
    Call BuildWorkbook
    gSilentBuild = False

    Exit Sub
Trap:
    gSilentBuild = False
    Dim n As Long, d As String, src As String
    n = Err.Number: d = Err.Description: src = Err.Source
    Dim fnum As Integer
    fnum = FreeFile
    Open logPath For Output As #fnum
    Print #fnum, "Err.Number=" & n
    Print #fnum, "Err.Source=" & src
    Print #fnum, "Err.Description=" & d
    Close #fnum
    Err.Raise n, src, d
End Sub
