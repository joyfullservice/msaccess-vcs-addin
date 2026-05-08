Attribute VB_Name = "modTestSuite"
Option Compare Database
Option Explicit
Option Private Module

Private Const ModuleName As String = "modTestSuite"

'@Folder("Tests")


' Encoding tests (TestUCS2toUTF8RoundTrip, TestParseSpecialCharsInJson, TestStringFileHash)
' moved to clsTestEncoding for class-based setup/teardown of temp files.


Public Sub TestSortDictionaryByKeys()

    Dim dItems As Dictionary

    Set dItems = New Dictionary
    dItems.Add "C", "C"
    dItems.Add "A", "A"
    dItems.Add "B", "B"

    Set dItems = SortDictionaryByKeys(dItems)

    TestAssert dItems.Items(0) = "A"
    TestAssert dItems.Items(1) = "B"
    TestAssert dItems.Items(2) = "C"

End Sub


Private Sub TestQuickSort()

    Dim astr() As String
    Dim strResult As String

    astr = Split("u i a")

    QuickSort astr
    strResult = Join(astr, " ")
    TestAssert strResult = "a i u"

End Sub


Private Sub TestConcat(Optional ManualRunOnly)

    With New clsConcat
        .SelfTest
    End With

End Sub


Private Sub TestSanitizeConnectionString()

    ' Verify semicolon placement matches original
    TestAssert SanitizeConnectionString(";test;test;") = ";test;test;"
    TestAssert SanitizeConnectionString("test;test") = "test;test"
    TestAssert SanitizeConnectionString(";test;test") = ";test;test"
    TestAssert SanitizeConnectionString("test;test;") = "test;test;"
    TestAssert SanitizeConnectionString("test;test;") = "test;test;"
    TestAssert SanitizeConnectionString("test") = "test"
    TestAssert SanitizeConnectionString(vbNullString) = vbNullString

End Sub


Private Sub TestCloneDictionary()

    Dim dFruit As Dictionary
    Dim dApple As Dictionary
    Dim dClone As Dictionary

    Set dFruit = New Dictionary
    Set dApple = New Dictionary

    ' Create text compare dictionary
    With dApple
        .CompareMode = TextCompare
        .Add "SEED1", "Apple Seed"
        .Add "seed2", "Apple Seed"
    End With

    ' Create binary compare dictionary with nested dictionary
    With dFruit
        .CompareMode = BinaryCompare
        .Add "Apple", dApple
        .Add "Orange", "Orange"
        .Add "Pear", "Pear"
    End With

    ' Clone the dictionary
    Set dClone = CloneDictionary(dFruit, ecmSourceMethod)

    ' Change some data in the cloned dictionary
    dClone("Apple")("Seed2") = "Pear Seed"

    ' Test the results to make sure it cloned correctly.
    TestAssert dClone.Exists("APPLE") = False
    TestAssert dClone.Exists("Apple") = True
    TestAssert dClone.Exists("ORANGE") = False
    TestAssert dClone.Exists("Orange") = True
    TestAssert dClone.CompareMode = BinaryCompare
    TestAssert dClone("Apple").CompareMode = Scripting.CompareMethod.TextCompare
    TestAssert dClone("Apple").Exists("seed1") = True
    TestAssert dClone("Apple").Exists("SEED1") = True
    TestAssert dClone("Apple").Exists("Seed3") = False
    TestAssert dClone("Apple")("Seed2") = "Pear Seed"
    TestAssert dFruit("Apple")("Seed2") = "Apple Seed"

End Sub


Private Sub TestComponentPropertyAccess()

    Dim cnt As IDbComponent
    Dim varTest As Variant

    For Each cnt In GetContainers
        ' Make sure none of the following throw an error
        ' when the database object has not been set.
        varTest = cnt.Name
        varTest = cnt.DateModified
        varTest = cnt.SourceFile
        TestAssert cnt.DbObject Is Nothing
    Next

End Sub


Private Sub TestUniqueComponentCategory()

    Dim dList As Dictionary
    Dim cnt As IDbComponent

    Set dList = New Dictionary
    For Each cnt In GetContainers
        TestAssert Not dList.Exists(cnt.Category)
        dList.Add cnt.Category, vbNullString
    Next

End Sub


Private Sub TestUniqueComponentType()

    Dim dList As Dictionary
    Dim cnt As IDbComponent

    Set dList = New Dictionary
    For Each cnt In GetContainers
        TestAssert Not dList.Exists(cnt.ComponentType)
        dList.Add cnt.ComponentType, vbNullString
    Next

End Sub


Private Sub TestUniqueBaseSubfolder()

    Dim dList As Dictionary
    Dim cnt As IDbComponent

    Set dList = New Dictionary
    For Each cnt In GetContainers
        If Not cnt.SingleFile Then
            TestAssert Not dList.Exists(cnt.BaseFolder)
            dList.Add cnt.BaseFolder, vbNullString
        End If
    Next

End Sub


' Test the operation of a progress bar without using label objects
' (Uses Access system progress meter instead)
Public Sub TestMeterProgressBar()

    Dim intCnt As Integer

    With New clsLblProg
        .Max = 20
        For intCnt = 1 To 30
            'Pause 0.1
            .Increment
        Next intCnt
        .Reset
        .Clear
    End With

End Sub


Public Sub TestGitRepositoryRoot()

    With New clsGitIntegration

        ' Verify repository root for this project
        TestAssert .GetRepositoryRoot = CurrentProject.Path & PathSep

        ' Resolve from subfolder
        .WorkingFolder = CurrentProject.Path & "\Version Control.accda.src\modules\"
        TestAssert .GetRepositoryRoot = CurrentProject.Path & PathSep

        ' Return working folder when not in a git repository
        ' (Also tests returning final path separator)
        .WorkingFolder = "c:\windows"
        TestAssert .GetRepositoryRoot = "c:\windows\"

        ' Reflect change in working folder
        .WorkingFolder = vbNullString
         TestAssert .GetRepositoryRoot = CurrentProject.Path & PathSep

        ' Return specified working folder, even if it doesn't exist
        .WorkingFolder = "c:\Some Path that Doesn't Exist"
         TestAssert .GetRepositoryRoot = "c:\Some Path that Doesn't Exist\"

    End With

End Sub


Public Sub TestInArray()
    Dim varArray As Variant
    varArray = Array("a", "b", "c", 1, 2, 3, #1/1/2000#)
    TestAssert InArray(varArray, "b")
    TestAssert Not InArray(varArray, "B")
    TestAssert InArray(varArray, "B", vbTextCompare)
    TestAssert InArray(varArray, 2)
    TestAssert InArray(varArray, #1/1/2000#)
    TestAssert Not InArray(varArray, Null)
    TestAssert Not InArray(Null, "b")
    TestAssert Not InArray(Array(), "b")
End Sub


Public Sub TestGetClassFromComponentType()

    Dim intType As eDatabaseComponentType

    ' Test the entire enum range of component types
    ' to make sure they are all assigned to a class.
    For intType = edbTableDataMacro To eDatabaseComponentType.[_Last] - 1
        TestAssert Not GetComponentClass(intType) Is Nothing
    Next intType

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestJsonNewLineIssue
' Author    : Adam Waller
' Date      : 7/24/2023
' Purpose   : Encountered an issue where vbCrLf strings are not parsed correctly when
'           : converting to JSON and back to string values.
'---------------------------------------------------------------------------------------
'
Public Sub TestJsonNewLineIssue()

    Const cstrTest As String = "Line1" & vbCrLf & "Line2" & vbCr & "Line3" & vbLf & "Line4" & vbCrLf

    Dim dTest As Dictionary
    Dim strResult As String

    Set dTest = New Dictionary

    dTest("Multiline") = cstrTest
    TestAssert dTest("Multiline") = cstrTest

    ' Test round trip conversion
    strResult = ParseJson(ConvertToJson(dTest, 2))("Multiline")
    TestAssert (strResult = cstrTest)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestSqlFormatter
' Author    : Adam Waller
' Date      : 8/16/2023
' Purpose   : Self-test the SQL Formatter class
'---------------------------------------------------------------------------------------
'
Public Sub TestSqlFormatter(Optional ManualRunOnly)
    With New clsSqlFormatter
        .SelfTest
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestCatch
' Author    : hecon5
' Date      : 10/20/2023
' Purpose   : Validates that Catch operates correctly and that LogUnhandledErrors
'           : doesn't create an infinite loop whether or not log exists.
'           :
'           : To use, run normally, after loading options / other core dependancies.
'           : Then Stop the code (in VBA IDE) and then run again. Stopping code execution
'---------------------------------------------------------------------------------------
'
Public Sub TestCatch()

    ' Specifiying a Const FunctionName allows copy/paste code and having the wrong FunctionName
    ' names if (when) they change.
    Const FunctionName As String = ModuleName & ".CatchTest"

    ' Make sure we don't trigger break mode
    Options.BreakOnError = False
    Operation.InteractionMode = eimSilent

    On Error Resume Next ' Clear out any errors that may happen, and continue on when errors happen.
    Err.Raise 24601, "Pre Log Test"

    ' This is the "standard" way of catching errors without losing them.
    LogUnhandledErrors FunctionName
    On Error Resume Next

    ' "Pretend" code tossing an error.
    Err.Raise 24602, "Post Log Test"
    ' Checking for any issues post code execution.
    CatchAny eelError, "Catch Test Validation", FunctionName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestPathFunctions
' Author    : Adam Waller
' Date      : 4/12/2025
' Purpose   : Ensure that VerifyPath is working correctly for different types of paths
'---------------------------------------------------------------------------------------
'
Public Sub TestPathFunctions()

    ' This path may not work on all systems, but it should work in a normal dev environment
    Const cstrUncBase As String = "\\%computername%\c$\users\%username%\AppData\Local\Temp\"

    Dim strBase As String
    Dim strPath As String
    Dim strTempPath As String
    Dim intCnt As Integer

    ' Test expansion of environment variable
    strPath = ExpandEnvironmentVariables("%TEMP%\test.tmp")
    TestAssert FSO.FolderExists(FSO.GetParentFolderName(strPath))

    ' Test relative path
    ' NOTE: strBase intentionally has NO trailing separator so callers below can
    ' concatenate "\sub..." segments without producing an empty path component
    ' (i.e. "Temp\\sub..."). Short paths get auto-normalized by SHCreateDirectoryEx,
    ' but the long-path branch uses the "\\?\" prefix which disables normalization
    ' and would reject "\\" with ERROR_INVALID_NAME.
    strBase = ExpandEnvironmentVariables("%TEMP%")
    strTempPath = strBase & "\subfolder\level2\"
    If FSO.FolderExists(strTempPath) Then FSO.DeleteFolder StripSlash(strTempPath)
    TestAssert Not FSO.FolderExists(strTempPath)
    TestAssert VerifyPath(strTempPath)
    TestAssert FSO.FolderExists(strTempPath)
    TestAssert GetRelativePath(strTempPath, strBase) = "rel:\subfolder\level2\"
    FSO.DeleteFolder strBase & "\subfolder"

    ' Test verify path with file name
    strTempPath = strTempPath & "test.tmp"
    TestAssert VerifyPath(strTempPath)
    TestAssert FSO.FolderExists(FSO.GetParentFolderName(strTempPath))
    FSO.DeleteFolder strBase & "\subfolder"

    ' Test UNC path (May not work on all systems)
    strTempPath = ExpandEnvironmentVariables(cstrUncBase & "subfolder\level2\test.tmp")
    TestAssert VerifyPath(strTempPath)
    TestAssert FSO.FolderExists(FSO.GetParentFolderName(strTempPath))
    FSO.DeleteFolder strBase & "\subfolder"

    ' BuildPath2 must preserve the leading "\\" UNC prefix on the first segment.
    ' Regression: a previous slash-stripping change collapsed "\\server\share\..."
    ' to "server\share\..." which made VerifyPath / SHCreateDirectoryEx report the
    ' path as relative (#issue: command bar image export against UNC export folder).
    TestAssert BuildPath2("\\server\share\root\", "menus", "name_Images") = _
        "\\server\share\root\menus\name_Images"
    TestAssert BuildPath2("\\server\share\root", "sub\") = _
        "\\server\share\root\sub"

    ' Non-UNC behaviour must remain unchanged: redundant separators between
    ' segments are still trimmed and a leading slash on a non-first segment is
    ' still stripped (e.g. BuildPath2(CurrentProject.Path, "\Template\..."))
    TestAssert BuildPath2("C:\foo\", "\bar\", "baz") = "C:\foo\bar\baz"
    TestAssert BuildPath2("C:\foo", "\Template\CommandBars.bin") = _
        "C:\foo\Template\CommandBars.bin"

    ' LONG PATHS (> 260) (Requires OS support and newer version of Access)
    'https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation?tabs=registry
    ' Gated on the OS-level LongPathsEnabled flag. Without it, Win32/shell APIs
    ' return ERROR_FILENAME_EXCED_RANGE (206) for any path > MAX_PATH regardless of
    ' a "\\?\" prefix, which would surface here as a false test failure.
    If Application.Version >= 16 And LongPathsEnabled() Then

        ' Test long path (On newer versions of Access)
        strTempPath = strBase & "\" & Repeat("subfolder\", 26)
        TestAssert VerifyPath(strTempPath)
        strPath = strBase & "\subfolder"
        If FSO.FolderExists(strPath) Then FSO.DeleteFolder strPath

        ' Test long UNC path
        strTempPath = cstrUncBase & Repeat("subfolder\", 26)
        TestAssert VerifyPath(strTempPath)
        strPath = strBase & "\subfolder"
        If FSO.FolderExists(strPath) Then FSO.DeleteFolder strPath
    ElseIf Application.Version >= 16 Then
        Debug.Print "TestPathFunctions: skipping long-path checks " & _
            "(HKLM\SYSTEM\CurrentControlSet\Control\FileSystem\LongPathsEnabled is not set)."
    End If

End Sub
