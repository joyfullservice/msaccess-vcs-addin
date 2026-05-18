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
