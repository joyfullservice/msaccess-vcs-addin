Attribute VB_Name = "modUnitTesting"
Option Compare Database
Option Explicit
Option Private Module

Private Const ModuleName As String = "modUnitTesting"

'@Folder("Tests")


' Test shows that UCS-2 files exported by Access make round trip through our conversions.
Public Sub TestUCS2toUTF8RoundTrip()

    Dim queryName As String
    queryName = "Temp_Test_Query_Delete_Me_Æ_ø_Å"
    Dim tempFileName As String
    tempFileName = GetTempFile()

    Dim UCStoUCS As String
    Dim UCStoUTF As String
    Dim UTFtoUTF As String
    Dim UTFtoUCS As String
    UCStoUCS = tempFileName & "UCS-2toUCS-2"
    UCStoUTF = tempFileName & "UCS-2toUTF-8"
    UTFtoUTF = tempFileName & "UTF-8toUTF-8"
    UTFtoUCS = tempFileName & "UTF-8toUCS-2"

    ' Use temporary query to export example file
    CurrentDb.CreateQueryDef queryName, "SELECT * FROM TEST WHERE TESTING='ÆØÅ'"
    Application.SaveAsText acQuery, queryName, tempFileName
    CurrentDb.QueryDefs.Delete queryName

    ' Read original export
    Dim originalExport As String
    With FSO.OpenTextFile(tempFileName, ForReading, False, TristateTrue)
        originalExport = .ReadAll
        .Close
    End With

    ConvertUtf8Ucs2 tempFileName, UCStoUCS
    ConvertUcs2Utf8 UCStoUCS, UCStoUTF
    ConvertUcs2Utf8 UCStoUTF, UTFtoUTF
    ConvertUtf8Ucs2 UTFtoUTF, UTFtoUCS

    ' Read final file that went through all permutations of conversion
    Dim finalFile As String
    With FSO.OpenTextFile(UTFtoUCS, ForReading, False, TristateTrue)
        finalFile = .ReadAll
        .Close
    End With

    Debug.Assert originalExport = finalFile

End Sub


Public Sub TestParseSpecialCharsInJson()

    Dim strPath As String
    Dim dict As Dictionary
    Dim FSO As Object

    strPath = GetTempFile

    Set FSO = CreateObject("Scripting.FileSystemObject")
    With FSO.CreateTextFile(strPath, True)
        .WriteLine "{""Test"":""ÆØÅ are special?""}"
        .Close
    End With

    Debug.Print strPath

    Set dict = modFileAccess.ReadJsonFile(strPath)

    Debug.Assert Not dict Is Nothing
    Debug.Print dict("Test")

End Sub


Public Sub TestSortDictionaryByKeys()

    Dim dItems As Dictionary

    Set dItems = New Dictionary
    dItems.Add "C", "C"
    dItems.Add "A", "A"
    dItems.Add "B", "B"

    Set dItems = SortDictionaryByKeys(dItems)

    Debug.Assert dItems.Items(0) = "A"
    Debug.Assert dItems.Items(1) = "B"
    Debug.Assert dItems.Items(2) = "C"

End Sub


Private Sub TestQuickSort()

    Dim astr() As String
    Dim strResult As String

    astr = Split("u i a")

    QuickSort astr
    strResult = Join(astr, " ")
    Debug.Assert strResult = "a i u"

End Sub


Private Sub TestConcat()

    With New clsConcat
        .SelfTest
    End With

End Sub


Private Sub TestSanitizeConnectionString()

    ' Verify semicolon placement matches original
    Debug.Assert SanitizeConnectionString(";test;test;") = ";test;test;"
    Debug.Assert SanitizeConnectionString("test;test") = "test;test"
    Debug.Assert SanitizeConnectionString(";test;test") = ";test;test"
    Debug.Assert SanitizeConnectionString("test;test;") = "test;test;"
    Debug.Assert SanitizeConnectionString("test;test;") = "test;test;"
    Debug.Assert SanitizeConnectionString("test") = "test"
    Debug.Assert SanitizeConnectionString(vbNullString) = vbNullString

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
    Debug.Assert dClone.Exists("APPLE") = False
    Debug.Assert dClone.Exists("Apple") = True
    Debug.Assert dClone.Exists("ORANGE") = False
    Debug.Assert dClone.Exists("Orange") = True
    Debug.Assert dClone.CompareMode = BinaryCompare
    Debug.Assert dClone("Apple").CompareMode = Scripting.CompareMethod.TextCompare
    Debug.Assert dClone("Apple").Exists("seed1") = True
    Debug.Assert dClone("Apple").Exists("SEED1") = True
    Debug.Assert dClone("Apple").Exists("Seed3") = False
    Debug.Assert dClone("Apple")("Seed2") = "Pear Seed"
    Debug.Assert dFruit("Apple")("Seed2") = "Apple Seed"

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
        Debug.Assert cnt.DbObject Is Nothing
    Next

End Sub


Private Sub TestUniqueComponentCategory()

    Dim dList As Dictionary
    Dim cnt As IDbComponent

    Set dList = New Dictionary
    For Each cnt In GetContainers
        Debug.Assert Not dList.Exists(cnt.Category)
        dList.Add cnt.Category, vbNullString
    Next

End Sub


Private Sub TestUniqueComponentType()

    Dim dList As Dictionary
    Dim cnt As IDbComponent

    Set dList = New Dictionary
    For Each cnt In GetContainers
        Debug.Assert Not dList.Exists(cnt.ComponentType)
        dList.Add cnt.ComponentType, vbNullString
    Next

End Sub


Private Sub TestUniqueBaseSubfolder()

    Dim dList As Dictionary
    Dim cnt As IDbComponent

    Set dList = New Dictionary
    For Each cnt In GetContainers
        If Not cnt.SingleFile Then
            Debug.Assert Not dList.Exists(cnt.BaseFolder)
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
        Debug.Assert .GetRepositoryRoot = CurrentProject.Path & PathSep

        ' Resolve from subfolder
        .WorkingFolder = CurrentProject.Path & "\Version Control.accda.src\modules\"
        Debug.Assert .GetRepositoryRoot = CurrentProject.Path & PathSep

        ' Return working folder when not in a git repository
        ' (Also tests returning final path separator)
        .WorkingFolder = "c:\windows"
        Debug.Assert .GetRepositoryRoot = "c:\windows\"

        ' Reflect change in working folder
        .WorkingFolder = vbNullString
         Debug.Assert .GetRepositoryRoot = CurrentProject.Path & PathSep

        ' Return specified working folder, even if it doesn't exist
        .WorkingFolder = "c:\Some Path that Doesn't Exist"
         Debug.Assert .GetRepositoryRoot = "c:\Some Path that Doesn't Exist\"

    End With

End Sub


Public Sub TestInArray()
    Dim varArray As Variant
    varArray = Array("a", "b", "c", 1, 2, 3, #1/1/2000#)
    Debug.Assert InArray(varArray, "b")
    Debug.Assert Not InArray(varArray, "B")
    Debug.Assert InArray(varArray, "B", vbTextCompare)
    Debug.Assert InArray(varArray, 2)
    Debug.Assert InArray(varArray, #1/1/2000#)
    Debug.Assert Not InArray(varArray, Null)
    Debug.Assert Not InArray(Null, "b")
    Debug.Assert Not InArray(Array(), "b")
End Sub


Public Sub TestStringFileHash()

    Const cstrText As String = "This is my text content."
    Dim strTempFile As String

    ' Make sure we get the same result when hashing a string as hashing a file.

    ' Create a file, and write our content.
    strTempFile = GetTempFile
    WriteFile cstrText, strTempFile

    ' Compare to known hash (without BOM)
    Debug.Assert GetStringHash(cstrText) = "f80a555"        ' Without BOM
    Debug.Assert GetStringHash(cstrText, True) = "b628391"  ' With UTF-8 BOM and trailing vbCrLf

    ' Compare results of hashing file with hashing a string.
    Debug.Assert GetFileHash(strTempFile) = GetStringHash(cstrText, True)

    ' Remove temp file.
    FSO.DeleteFile strTempFile

End Sub


Public Sub TestGetClassFromComponentType()

    Dim intType As eDatabaseComponentType

    ' Test the entire enum range of component types
    ' to make sure they are all assigned to a class.
    For intType = edbTableDataMacro To eDatabaseComponentType.[_Last] - 1
        Debug.Assert Not GetComponentClass(intType) Is Nothing
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
    Debug.Assert dTest("Multiline") = cstrTest

    ' Test round trip conversion
    strResult = ParseJson(ConvertToJson(dTest, 2))("Multiline")
    Debug.Assert (strResult = cstrTest)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestSqlFormatter
' Author    : Adam Waller
' Date      : 8/16/2023
' Purpose   : Self-test the SQL Formatter class
'---------------------------------------------------------------------------------------
'
Public Sub TestSqlFormatter()
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
    Debug.Assert FSO.FolderExists(FSO.GetParentFolderName(strPath))

    ' Test relative path
    strBase = ExpandEnvironmentVariables("%TEMP%\")
    strTempPath = strBase & "\subfolder\level2\"
    If FSO.FolderExists(strTempPath) Then FSO.DeleteFolder StripSlash(strTempPath)
    Debug.Assert Not FSO.FolderExists(strTempPath)
    Debug.Assert VerifyPath(strTempPath)
    Debug.Assert FSO.FolderExists(strTempPath)
    Debug.Assert GetRelativePath(strTempPath, strBase) = "rel:\subfolder\level2\"
    FSO.DeleteFolder strBase & "\subfolder"

    ' Test verify path with file name
    strTempPath = strTempPath & "test.tmp"
    Debug.Assert VerifyPath(strTempPath)
    Debug.Assert FSO.FolderExists(FSO.GetParentFolderName(strTempPath))
    FSO.DeleteFolder strBase & "\subfolder"

    ' Test UNC path (May not work on all systems)
    strTempPath = ExpandEnvironmentVariables(cstrUncBase & "subfolder\level2\test.tmp")
    Debug.Assert VerifyPath(strTempPath)
    Debug.Assert FSO.FolderExists(FSO.GetParentFolderName(strTempPath))
    FSO.DeleteFolder strBase & "\subfolder"

    ' LONG PATHS (> 260) (Requires OS support and newer version of Access)
    'https://learn.microsoft.com/en-us/windows/win32/fileio/maximum-file-path-limitation?tabs=registry
    If Application.Version >= 16 Then

        ' Test long path (On newer versions of Access)
        strTempPath = strBase & "\" & Repeat("subfolder\", 26)
        Debug.Assert VerifyPath(strTempPath)
        strPath = strBase & "\subfolder"
        If FSO.FolderExists(strPath) Then FSO.DeleteFolder strPath

        ' Test long UNC path
        strTempPath = cstrUncBase & Repeat("subfolder\", 26)
        Debug.Assert VerifyPath(strTempPath)
        strPath = strBase & "\subfolder"
        If FSO.FolderExists(strPath) Then FSO.DeleteFolder strPath
    End If

End Sub
