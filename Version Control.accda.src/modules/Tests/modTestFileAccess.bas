Attribute VB_Name = "modTestFileAccess"
'---------------------------------------------------------------------------------------
' Module    : modTestFileAccess
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Tests for modFileAccess path and file functions.
'           : Migrated TestPathFunctions from modTestSuite.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.FileIO")


Public Sub TestPathFunctions()

    Const cstrUncBase As String = "\\%computername%\c$\users\%username%\AppData\Local\Temp\"

    Dim strBase As String
    Dim strPath As String
    Dim strTempPath As String

    ' Test expansion of environment variable
    strPath = ExpandEnvironmentVariables("%TEMP%\test.tmp")
    TestAssert FSO.FolderExists(FSO.GetParentFolderName(strPath)), "TEMP folder exists"

    ' Test relative path
    strBase = ExpandEnvironmentVariables("%TEMP%")
    strTempPath = strBase & "\subfolder\level2\"
    If FSO.FolderExists(strTempPath) Then FSO.DeleteFolder StripSlash(strTempPath)
    TestAssert Not FSO.FolderExists(strTempPath), "temp path doesn't exist yet"
    TestAssert VerifyPath(strTempPath), "VerifyPath creates folders"
    TestAssert FSO.FolderExists(strTempPath), "temp path now exists"
    TestAssert GetRelativePath(strTempPath, strBase) = "rel:\subfolder\level2\", "relative path"
    FSO.DeleteFolder strBase & "\subfolder"

    ' Test verify path with file name
    strTempPath = strBase & "\subfolder\level2\test.tmp"
    TestAssert VerifyPath(strTempPath), "VerifyPath with file name"
    TestAssert FSO.FolderExists(FSO.GetParentFolderName(strTempPath)), "parent folder created"
    FSO.DeleteFolder strBase & "\subfolder"

    ' Test UNC path (may not work on all systems)
    strTempPath = ExpandEnvironmentVariables(cstrUncBase & "subfolder\level2\test.tmp")
    TestAssert VerifyPath(strTempPath), "UNC path"
    TestAssert FSO.FolderExists(FSO.GetParentFolderName(strTempPath)), "UNC folder created"
    FSO.DeleteFolder strBase & "\subfolder"

End Sub


Public Sub TestBuildPath2()
    TestAssert BuildPath2("\\server\share\root\", "menus", "name_Images") = _
        "\\server\share\root\menus\name_Images", "UNC prefix preserved"
    TestAssert BuildPath2("\\server\share\root", "sub\") = _
        "\\server\share\root\sub", "trailing slash trimmed"
    TestAssert BuildPath2("C:\foo\", "\bar\", "baz") = "C:\foo\bar\baz", _
        "redundant separators trimmed"
    TestAssert BuildPath2("C:\foo", "\Template\CommandBars.bin") = _
        "C:\foo\Template\CommandBars.bin", "leading slash on second segment"
End Sub


Public Sub TestLongPaths()
    ' Long paths require OS support and Access 2016+
    If Application.Version < 16 Or Not LongPathsEnabled() Then Exit Sub

    Dim strBase As String
    Dim strTempPath As String
    Dim strPath As String

    strBase = ExpandEnvironmentVariables("%TEMP%")
    strTempPath = strBase & "\" & Repeat("subfolder\", 26)
    TestAssert VerifyPath(strTempPath), "long path created"
    strPath = strBase & "\subfolder"
    If FSO.FolderExists(strPath) Then FSO.DeleteFolder strPath
End Sub
