Attribute VB_Name = "modTestSuite"
'---------------------------------------------------------------------------------------
' Module    : modTestSuite
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Legacy test suite. Most tests have been migrated to focused test modules:
'           :   modTestCollectionUtil, modTestFunctions, modTestConnect,
'           :   modTestJsonConverter, clsTestSqlFormatter, modTestFileAccess,
'           :   modTestComponentInvariants, modTestErrorHandling.
'           : Remaining tests below are either unique to this module or not yet migrated.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests")
'@Tag("integration")


Public Sub TestMeterProgressBar()

    Dim intCnt As Integer

    With New clsLblProg
        .Max = 20
        For intCnt = 1 To 30
            .Increment
        Next intCnt
        TestAssert .Value = 20, "value clamped to Max after overflow"
        .Reset
        TestAssert .Value = 0, "value reset to zero"
        .Clear
    End With

End Sub


Public Sub TestGitRepositoryRoot()

    With New clsGitIntegration

        ' Verify repository root for this project
        TestAssert .GetRepositoryRoot = CurrentProject.Path & PathSep, "current project path"

        ' Resolve from subfolder
        .WorkingFolder = CurrentProject.Path & "\Version Control.accda.src\modules\"
        TestAssert .GetRepositoryRoot = CurrentProject.Path & PathSep, "from subfolder"

        ' Return working folder when not in a git repository
        .WorkingFolder = "c:\windows"
        TestAssert .GetRepositoryRoot = "c:\windows\", "non-repo folder"

        ' Reflect change in working folder
        .WorkingFolder = vbNullString
        TestAssert .GetRepositoryRoot = CurrentProject.Path & PathSep, "reset to default"

        ' Return specified working folder, even if it doesn't exist
        .WorkingFolder = "c:\Some Path that Doesn't Exist"
        TestAssert .GetRepositoryRoot = "c:\Some Path that Doesn't Exist\", "non-existent path"

    End With

End Sub


Public Sub TestGetClassFromComponentType()

    Dim intType As eDatabaseComponentType

    ' Every enum value should map to a class
    For intType = edbTableDataMacro To eDatabaseComponentType.[_Last] - 1
        TestAssert Not GetComponentClass(intType) Is Nothing, _
            "component type " & intType & " has a class"
    Next intType

End Sub


Public Sub TestSidecarClassGating()

    Dim lngSaved As Long
    Dim strJsonFile As String
    Dim dFile As Dictionary
    Dim cComp As IDbComponent
    Const strModuleName As String = "modTestAssert"

    ' TypeName(cComp) drives Info.Class; an unbound instance is sufficient here.
    Set cComp = New clsDbModule

    lngSaved = Options.ExportFormatVersion
    strJsonFile = GetTempFolder("vcs_sidecar_class") & PathSep & strModuleName & ".json"
    If FSO.FileExists(strJsonFile) Then DeleteFile strJsonFile

    ' Seed a companion file without Info so ExportObjectMetadata creates the header.
    ' A non-metadata Items key ensures the file is retained even when DAO has no Description.
    WriteFile "{""Items"":{""TestRetention"":true}}", strJsonFile

    ' Below 5.1.0 the owning class name is ignored and Info.Class stays empty.
    Options.ExportFormatVersion = EFV_5_0_0
    ExportObjectMetadata strJsonFile, "Modules", strModuleName, cComp
    Set dFile = ReadJsonFile(strJsonFile)
    TestAssert Not dFile Is Nothing, "companion file written at 5.0.0"
    TestAssert dFile.Exists("Info"), "Info header created at 5.0.0"
    TestAssert dFile("Info")("Class") = vbNullString, "5.0.0 uses empty Class on new companion"

    ' At 5.1.0 the owning component class (via TypeName) is recorded and healed in.
    Options.ExportFormatVersion = EFV_5_1_0
    ExportObjectMetadata strJsonFile, "Modules", strModuleName, cComp
    Set dFile = ReadJsonFile(strJsonFile)
    TestAssert dFile("Info")("Class") = "clsDbModule", "5.1.0 sets owning component class"

    If FSO.FileExists(strJsonFile) Then DeleteFile strJsonFile
    Options.ExportFormatVersion = lngSaved

End Sub
