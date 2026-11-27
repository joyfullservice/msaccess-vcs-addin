Attribute VB_Name = "modTestSourceFolderCopy"
'---------------------------------------------------------------------------------------
' Module    : modTestSourceFolderCopy
' Author    : Adam Waller
' Date      : 11/27/2026
' Purpose   : Unit tests for IsSourceFolderCopy and the build path it decides: a copy
'           : of the source (a git worktree or clone) of a project with an absolute
'           : ExportFolder builds beside itself, not at the original's VCS Build Path.
'           : Uses small json files in temp folders; no real build.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")
'@Tag("unit")


Public Sub TestCopyOfAbsoluteSourceIsDetected()

    Dim strRoot As String
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    strRoot = GetTempFolder("vcs_source_copy")
    WriteSourceFolder strRoot & "\orig\src", strRoot & "\orig\src\"
    WriteSourceFolder strRoot & "\copy\src", strRoot & "\orig\src\"

    TestAssert Not IsSourceFolderCopy(strRoot & "\orig\src"), "the configured folder is not a copy"
    TestAssert IsSourceFolderCopy(strRoot & "\copy\src"), "another folder with the same absolute setting is a copy"
    TestAssert IsSourceFolderCopy(strRoot & "\copy\src\"), "a trailing separator does not matter"

CleanUp:
    On Error Resume Next
    If Len(strRoot) Then FSO.DeleteFolder strRoot, True
    If lngErr <> 0 Then TestAssert False, "copy detection " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestRelativeOrMissingSettingIsNeverACopy()

    Dim strRoot As String
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    strRoot = GetTempFolder("vcs_source_copy")
    WriteSourceFolder strRoot & "\blank\src", vbNullString
    WriteSourceFolder strRoot & "\relative\src", "\src\"
    WriteSourceFolder strRoot & "\Example.accdb", strRoot & "\%dbName%\"
    FSO.CreateFolder strRoot & "\missing"

    TestAssert Not IsSourceFolderCopy(strRoot & "\blank\src"), "blank resolves beside each database"
    TestAssert Not IsSourceFolderCopy(strRoot & "\relative\src"), "a relative setting resolves beside each database"
    TestAssert Not IsSourceFolderCopy(strRoot & "\missing"), "no vcs-options.json means the default"
    TestAssert Not IsSourceFolderCopy(strRoot & "\Example.accdb"), "%dbName% expands to the file in vbe-project.json"

CleanUp:
    On Error Resume Next
    If Len(strRoot) Then FSO.DeleteFolder strRoot, True
    If lngErr <> 0 Then TestAssert False, "relative settings " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestBuildPathForCopyAndOriginal()

    Dim strRoot As String
    Dim strSaved As String
    Dim blnSaved As Boolean
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    strRoot = GetTempFolder("vcs_source_copy")
    WriteSourceFolder strRoot & "\orig\src", strRoot & "\orig\src\", strRoot & "\apps"
    WriteSourceFolder strRoot & "\copy\src", strRoot & "\orig\src\", strRoot & "\apps"

    ' A build sets Options.ExportFolder to the folder it reads from (modBuild.Build).
    strSaved = Options.ExportFolder
    blnSaved = True

    Options.ExportFolder = strRoot & "\copy\src"
    TestAssert StrComp(GetOriginalDbFullPathFromSource(strRoot & "\copy\src"), _
        strRoot & "\copy\Example.accdb", vbTextCompare) = 0, "a copy builds beside itself"

    Options.ExportFolder = strRoot & "\orig\src"
    TestAssert StrComp(GetOriginalDbFullPathFromSource(strRoot & "\orig\src"), _
        strRoot & "\apps\Example.accdb", vbTextCompare) = 0, "the configured folder still builds at VCS Build Path"

CleanUp:
    On Error Resume Next
    If blnSaved Then Options.ExportFolder = strSaved
    If Len(strRoot) Then FSO.DeleteFolder strRoot, True
    If lngErr <> 0 Then TestAssert False, "build path " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestGitWorkingTreeRoot()

    Dim strRoot As String
    Dim lngErr As Long
    Dim strErr As String

    On Error GoTo ErrHandler

    strRoot = GetTempFolder("vcs_working_tree")
    FSO.CreateFolder strRoot & "\clone"
    FSO.CreateFolder strRoot & "\clone\.git"
    FSO.CreateFolder strRoot & "\clone\app"
    FSO.CreateFolder strRoot & "\worktree"
    FSO.CreateFolder strRoot & "\worktree\app"
    WriteFile "gitdir: C:\example\.git\worktrees\worktree", strRoot & "\worktree\.git"
    FSO.CreateFolder strRoot & "\plain"

    TestAssert StrComp(GetGitWorkingTreeRoot(strRoot & "\clone\app"), strRoot & "\clone", vbTextCompare) = 0, _
        "a .git folder marks a checkout or clone"
    TestAssert StrComp(GetGitWorkingTreeRoot(strRoot & "\worktree\app"), strRoot & "\worktree", vbTextCompare) = 0, _
        "a .git file marks a linked worktree"
    TestAssert Len(GetGitWorkingTreeRoot(strRoot & "\plain")) = 0 _
        Or Not IsPathInFolder(GetGitWorkingTreeRoot(strRoot & "\plain"), strRoot), _
        "a folder outside any working tree has none of its own"

CleanUp:
    On Error Resume Next
    If Len(strRoot) Then FSO.DeleteFolder strRoot, True
    If lngErr <> 0 Then TestAssert False, "working tree root " & lngErr & ": " & strErr
    Exit Sub

ErrHandler:
    lngErr = Err.Number
    strErr = Err.Description
    Resume CleanUp

End Sub


Public Sub TestIsPathInFolder()
    TestAssert IsPathInFolder("C:\Example\wt\src", "C:\example\WT\"), "a subfolder is inside"
    TestAssert IsPathInFolder("C:\Example\wt", "C:\Example\wt"), "a folder is inside itself"
    TestAssert Not IsPathInFolder("C:\Example\wt-other\src", "C:\Example\wt"), "a sibling with the same prefix is outside"
    TestAssert Not IsPathInFolder("C:\Example\wt\..\main\src", "C:\Example\wt"), "dot segments are resolved first"
End Sub


Public Sub TestFolderPathsMatch()
    TestAssert FolderPathsMatch("C:\Example\src", "c:\example\SRC\"), "case and trailing separator are ignored"
    TestAssert FolderPathsMatch("C:\Example\other\..\src", "C:\Example\src"), "dot segments are resolved"
    TestAssert Not FolderPathsMatch("C:\Example\src", "C:\Example\src2"), "a longer name is a different folder"
    TestAssert Not FolderPathsMatch(vbNullString, "C:\Example"), "an empty path matches nothing"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : WriteSourceFolder
' Purpose   : Create the minimum source files GetOriginalDbFullPathFromSource reads.
'---------------------------------------------------------------------------------------
'
Private Sub WriteSourceFolder(strFolder As String, strExportFolder As String, _
    Optional strBuildPath As String)

    VerifyPath strFolder & PathSep & "vbe-project.json"
    WriteFile "{""Items"":{""FileName"":""Example.accdb""}}", strFolder & PathSep & "vbe-project.json"
    WriteFile "{""Options"":{""ExportFolder"":""" & Replace(strExportFolder, "\", "\\") & """}}", _
        strFolder & PathSep & "vcs-options.json"
    If Len(strBuildPath) Then
        WriteFile "{""Items"":{""VCS Build Path"":""" & Replace(strBuildPath, "\", "\\") & """}}", _
            strFolder & PathSep & "proj-properties.json"
    End If

End Sub
