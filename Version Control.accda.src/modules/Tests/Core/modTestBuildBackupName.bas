Attribute VB_Name = "modTestBuildBackupName"
'---------------------------------------------------------------------------------------
' Module    : modTestBuildBackupName
' Author    : Adam Waller
' Date      : 11/19/2026
' Purpose   : Unit tests for GetBackupFileName, the name a build saves the original
'           : database under. Uses empty files in a temp folder; no real build.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")
'@Tag("unit")


Public Sub TestBackupNameSkipsNamesInUse()
    Dim strFolder As String

    strFolder = GetTempFolder("vcs_backup_name") & PathSep
    TestAssert GetBackupFileName(strFolder & "Example.accdb") = strFolder & "Example_VCSBackup.accdb", _
        "first backup has no number"

    FSO.CreateTextFile(strFolder & "Example_VCSBackup.accdb").Close
    FSO.CreateTextFile(strFolder & "Example_VCSBackup1.accdb").Close
    TestAssert GetBackupFileName(strFolder & "Example.accdb") = strFolder & "Example_VCSBackup2.accdb", _
        "next free number is used"

    FSO.DeleteFolder StripSlash(strFolder), True
End Sub


Public Sub TestBackupNameEmptyWhenEveryNameInUse()
    Dim strFolder As String
    Dim intCnt As Integer

    strFolder = GetTempFolder("vcs_backup_name") & PathSep
    FSO.CreateTextFile(strFolder & "Example_VCSBackup.accdb").Close
    For intCnt = 1 To 499
        FSO.CreateTextFile(strFolder & "Example_VCSBackup" & intCnt & ".accdb").Close
    Next intCnt
    TestAssert GetBackupFileName(strFolder & "Example.accdb") = vbNullString, _
        "no name is returned once all 500 are taken"

    FSO.DeleteFolder StripSlash(strFolder), True
End Sub
