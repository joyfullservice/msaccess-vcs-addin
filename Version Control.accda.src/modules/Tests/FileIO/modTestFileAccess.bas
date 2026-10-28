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
'@Tag("io")


Public Sub TestPathFunctions()

    Dim strBase As String
    Dim strPath As String
    Dim strTempPath As String
    Dim strUncBase As String

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

    ' Test UNC path (skip when admin share of TEMP is unreachable)
    strUncBase = LocalPathAsAdminShareUnc(strBase)
    If Len(strUncBase) > 0 And FSO.FolderExists(strUncBase) Then
        strTempPath = strUncBase & "\subfolder\level2\test.tmp"
        TestAssert VerifyPath(strTempPath), "UNC path"
        TestAssert FSO.FolderExists(FSO.GetParentFolderName(strTempPath)), "UNC folder created"
        FSO.DeleteFolder strUncBase & "\subfolder"
    End If

End Sub


Public Sub TestGetUncPathEnvironmentVariables()

    Dim strExpanded As String
    Dim strResult As String
    Const cstrMissingVar As String = "%NONEXISTENT_VCS_TEST_VAR%\foo"

    strExpanded = ExpandEnvironmentVariables("%TEMP%\foo")
    strResult = GetUncPath("%TEMP%\foo")
    TestAssert InStr(1, strResult, FSO.GetDriveName(strExpanded), vbTextCompare) > 0, _
        "GetUncPath expands %TEMP%"
    TestAssert GetUncPath(cstrMissingVar) = cstrMissingVar, _
        "GetUncPath returns unchanged path when env var unset"

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


Public Sub TestWriteFileSkipsUnchangedContent()

    Dim strPath As String
    Dim strContent As String
    Dim dteBefore As Date
    Dim dteAfter As Date

    strPath = ExpandEnvironmentVariables("%TEMP%\vcs_write_skip_test.txt")
    If FSO.FileExists(strPath) Then DeleteFile strPath

    strContent = "VCS WriteFile skip test" & vbCrLf
    WriteFile strContent, strPath
    TestAssert FSO.FileExists(strPath), "file created"
    TestAssert ReadFile(strPath) = strContent, "initial content matches"

    ' Backdate the file so an unchanged rewrite is distinguishable without waiting
    ' out the one-second filesystem timestamp resolution.
    SetFileDate strPath, DateAdd("n", -5, Now), True
    dteBefore = GetLastModifiedDate(strPath)
    WriteFile strContent, strPath
    dteAfter = GetLastModifiedDate(strPath)
    TestAssert dteBefore = dteAfter, "identical rewrite preserves DateLastModified"

    WriteFile strContent & "changed", strPath
    TestAssert ReadFile(strPath) <> strContent, "changed content written"
    TestAssert GetLastModifiedDate(strPath) > dteAfter, "changed rewrite updates DateLastModified"

    WriteFile vbNullString, strPath
    TestAssert Not FSO.FileExists(strPath), "empty string deletes file"

    If FSO.FileExists(strPath) Then DeleteFile strPath

End Sub


Public Sub TestWriteFileCaseCorrection()

    Dim strFolder As String
    Dim strPathLower As String
    Dim strPathUpper As String
    Dim dteBefore As Date
    Dim dteAfter As Date
    Dim strContent As String

    strFolder = ExpandEnvironmentVariables("%TEMP%\vcs_write_case_test")
    If FSO.FolderExists(strFolder) Then FSO.DeleteFolder strFolder, True
    VerifyPath strFolder & "\placeholder.txt"

    strPathLower = strFolder & "\casefile.txt"
    strPathUpper = strFolder & "\CASEFILE.txt"
    strContent = "case correction test" & vbCrLf

    WriteFile strContent, strPathLower
    TestAssert FSO.FileExists(strPathLower), "lowercase path file created"

    SetFileDate strPathLower, DateAdd("n", -5, Now), True
    dteBefore = GetLastModifiedDate(strPathLower)
    WriteFile strContent, strPathUpper
    TestAssert FSO.FileExists(strPathUpper), "uppercase path still resolves to file"
    dteAfter = GetLastModifiedDate(strPathUpper)
    TestAssert dteAfter > dteBefore, "case correction rewrite updates DateLastModified"

    FSO.DeleteFolder strFolder, True

End Sub


Public Sub TestWriteBinaryFileSkipsUnchangedContent()

    Dim strPath As String
    Dim bteOriginal(0 To 4) As Byte
    Dim bteChanged(0 To 4) As Byte
    Dim dteBefore As Date
    Dim dteAfter As Date
    Dim lngIdx As Long

    strPath = ExpandEnvironmentVariables("%TEMP%\vcs_write_binary_skip_test.bin")
    If FSO.FileExists(strPath) Then DeleteFile strPath

    For lngIdx = 0 To 4
        bteOriginal(lngIdx) = lngIdx + 1
        bteChanged(lngIdx) = lngIdx + 1
    Next lngIdx
    bteChanged(4) = 6

    WriteBinaryFile strPath, bteOriginal
    TestAssert FSO.FileExists(strPath), "binary file created"

    SetFileDate strPath, DateAdd("n", -5, Now), True
    dteBefore = GetLastModifiedDate(strPath)
    WriteBinaryFile strPath, bteOriginal
    dteAfter = GetLastModifiedDate(strPath)
    TestAssert dteBefore = dteAfter, "identical binary rewrite preserves DateLastModified"

    WriteBinaryFile strPath, bteChanged
    TestAssert GetBytesHash(GetFileBytes(strPath)) = GetBytesHash(bteChanged), "changed binary content written"
    TestAssert GetLastModifiedDate(strPath) > dteAfter, "changed binary rewrite updates DateLastModified"

    DeleteFile strPath

End Sub


Public Sub TestGetFileInfo()

    Dim strPath As String
    Dim dblSize As Double
    Dim strActualName As String

    strPath = ExpandEnvironmentVariables("%TEMP%\vcs_getfileinfo_test.txt")
    If FSO.FileExists(strPath) Then DeleteFile strPath

    TestAssert Not GetFileInfo(strPath, dblSize, strActualName), "missing file returns False"

    WriteFile "info test" & vbCrLf, strPath
    TestAssert GetFileInfo(strPath, dblSize, strActualName), "existing file returns True"
    TestAssert dblSize = FSO.GetFile(strPath).Size, "size matches FSO"
    TestAssert strActualName = FSO.GetFileName(strPath), "name matches path"

    DeleteFile strPath

End Sub


Public Sub TestDeleteFile()

    Dim strFolder As String
    Dim strPath1 As String
    Dim strPath2 As String
    Dim strPathOther As String
    Dim strMissing As String

    Const cstrExt As String = "vcs_del_test"

    strFolder = ExpandEnvironmentVariables("%TEMP%\vcs_delete_file_test")
    If FSO.FolderExists(strFolder) Then FSO.DeleteFolder strFolder, True
    VerifyPath strFolder & "\placeholder.txt"

    strPath1 = strFolder & "\alpha." & cstrExt
    strPath2 = strFolder & "\beta." & cstrExt
    strPathOther = strFolder & "\keep.txt"
    strMissing = strFolder & "\missing." & cstrExt

    WriteFile "one" & vbCrLf, strPath1
    WriteFile "two" & vbCrLf, strPath2
    WriteFile "other" & vbCrLf, strPathOther

    TestAssert FSO.FileExists(strPath1), "setup: first test file exists"
    TestAssert FSO.FileExists(strPath2), "setup: second test file exists"

    DeleteFile strMissing
    TestAssert Not FSO.FileExists(strMissing), "missing single file delete is no-op"

    DeleteFile strPathOther
    TestAssert Not FSO.FileExists(strPathOther), "single file delete removes file"
    DeleteFile strPathOther
    TestAssert Not FSO.FileExists(strPathOther), "repeat single file delete is no-op"

    DeleteFile strFolder & "\*." & cstrExt
    TestAssert Not FSO.FileExists(strPath1), "wildcard delete removes first file"
    TestAssert Not FSO.FileExists(strPath2), "wildcard delete removes second file"

    DeleteFile strFolder & "\*." & cstrExt

    WriteFile "again" & vbCrLf, strPath1
    WriteFile "other" & vbCrLf, strPathOther

    ClearFilesByExtension strFolder, cstrExt
    TestAssert Not FSO.FileExists(strPath1), "ClearFilesByExtension removes matching files"
    TestAssert FSO.FileExists(strPathOther), "ClearFilesByExtension preserves other extension"

    ClearFilesByExtension strFolder, cstrExt

    FSO.DeleteFolder strFolder, True

End Sub


Private Function LocalPathAsAdminShareUnc(ByVal strPath As String) As String

    strPath = StripSlash(strPath)
    If Left$(strPath, 2) = "\\" Then
        LocalPathAsAdminShareUnc = strPath
    ElseIf Mid$(strPath, 2, 1) = ":" Then
        LocalPathAsAdminShareUnc = "\\" & Environ$("COMPUTERNAME") & "\" _
            & Left$(strPath, 1) & "$" & Mid$(strPath, 3)
    End If

End Function
