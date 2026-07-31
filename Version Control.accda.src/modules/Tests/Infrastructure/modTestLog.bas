Attribute VB_Name = "modTestLog"
'---------------------------------------------------------------------------------------
' Module    : modTestLog
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Tests for clsLog path handling (SavedLogFilePath vs derived LogFilePath).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")


Public Sub TestLogSavedFilePath_DefaultSave()

    Dim strFolder As String
    Dim strSaved As String

    strFolder = GetTempFolder("vcs_log_saved_path") & PathSep
    Log.Clear
    Log.SourcePath = strFolder
    Log.Add "SavedLogFilePath regression", False
    Log.SaveFile

    strSaved = Log.SavedLogFilePath
    TestAssert Len(strSaved) > 0, "SavedLogFilePath set after SaveFile"
    TestAssert FSO.FileExists(strSaved), "SavedLogFilePath points at file on disk"
    TestAssert InStr(1, strSaved, "logs" & PathSep, vbTextCompare) > 0, _
        "default save uses logs subfolder"

    If FSO.FileExists(strSaved) Then DeleteFile strSaved

End Sub


Public Sub TestLogSavedFilePath_ClearWipesSavedPath()

    Dim strFolder As String

    strFolder = GetTempFolder("vcs_log_saved_path") & PathSep
    Log.Clear
    Log.SourcePath = strFolder
    Log.Add "before clear", False
    Log.SaveFile
    TestAssert Len(Log.SavedLogFilePath) > 0, "path recorded before Clear"

    Log.Clear
    TestAssert Log.SavedLogFilePath = vbNullString, "Clear resets SavedLogFilePath"

End Sub


Public Sub TestLogSavedFilePath_AlternatePath()

    Dim strFolder As String
    Dim strAlt As String

    strFolder = GetTempFolder("vcs_log_saved_path") & PathSep
    strAlt = FSO.BuildPath(strFolder, "logs" & PathSep & "CustomAlternate.log")

    Log.Clear
    Log.SourcePath = strFolder
    Log.Add "alternate path save", False
    Log.SaveFile strAlt

    TestAssert Log.SavedLogFilePath = strAlt, "SavedLogFilePath reflects alternate path"
    TestAssert FSO.FileExists(strAlt), "alternate path file exists"
    TestAssert StrComp(Log.SavedLogFilePath, Log.LogFilePath, vbTextCompare) <> 0, _
        "alternate save differs from derived LogFilePath"

    If FSO.FileExists(strAlt) Then DeleteFile strAlt

End Sub
