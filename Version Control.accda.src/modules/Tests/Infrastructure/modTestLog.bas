Attribute VB_Name = "modTestLog"
'---------------------------------------------------------------------------------------
' Module    : modTestLog
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Tests for clsLog path handling. SavedLogFilePath records the file that
'           : SaveFile actually wrote, where LogFilePath only derives a prospective
'           : name from state that moves during an operation.
'           : These tests use their own clsLog instance rather than the Log singleton,
'           : which is in use logging the test run itself.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")


Public Sub TestLogSavedFilePath_MatchesFileWritten()

    Dim cLog As clsLog
    Dim strSaved As String

    Set cLog = New clsLog
    cLog.SourcePath = GetTempFolder("vcs_log_saved_path") & PathSep
    cLog.Add "SavedLogFilePath regression", False

    TestAssert cLog.SavedLogFilePath = vbNullString, "no saved path before first SaveFile"

    cLog.SaveFile
    strSaved = cLog.SavedLogFilePath

    TestAssert Len(strSaved) > 0, "SavedLogFilePath set after SaveFile"
    TestAssert FSO.FileExists(strSaved), "SavedLogFilePath names a file on disk"

    If FSO.FileExists(strSaved) Then DeleteFile strSaved

End Sub


Public Sub TestLogSavedFilePath_ClearResetsPath()

    Dim cLog As clsLog
    Dim strSaved As String

    Set cLog = New clsLog
    cLog.SourcePath = GetTempFolder("vcs_log_saved_path") & PathSep
    cLog.Add "before clear", False
    cLog.SaveFile
    strSaved = cLog.SavedLogFilePath
    TestAssert Len(strSaved) > 0, "path recorded before Clear"

    ' Clear starts a fresh log (and a new operation ID), so the previously saved
    ' path no longer describes the current log content.
    cLog.Clear
    TestAssert cLog.SavedLogFilePath = vbNullString, "Clear resets SavedLogFilePath"

    If FSO.FileExists(strSaved) Then DeleteFile strSaved

End Sub


Public Sub TestLogSavedFilePath_AlternatePath()

    Dim cLog As clsLog
    Dim strFolder As String
    Dim strAlt As String

    Set cLog = New clsLog
    strFolder = GetTempFolder("vcs_log_saved_path") & PathSep
    strAlt = FSO.BuildPath(strFolder, "logs" & PathSep & "CustomAlternate.log")

    cLog.SourcePath = strFolder
    cLog.Add "alternate path save", False
    cLog.SaveFile strAlt

    TestAssert cLog.SavedLogFilePath = strAlt, "SavedLogFilePath reflects alternate path"
    TestAssert FSO.FileExists(strAlt), "alternate path file exists"
    TestAssert StrComp(cLog.SavedLogFilePath, cLog.LogFilePath, vbTextCompare) <> 0, _
        "alternate save differs from derived LogFilePath"

    If FSO.FileExists(strAlt) Then DeleteFile strAlt

End Sub
