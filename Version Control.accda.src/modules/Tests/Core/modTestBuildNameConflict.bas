Attribute VB_Name = "modTestBuildNameConflict"
'---------------------------------------------------------------------------------------
' Module    : modTestBuildNameConflict
' Author    : Adam Waller
' Date      : 8/25/2026
' Purpose   : Unit tests for ShouldCheckBuildNameConflict. Assertions on the helper
'           : only — no MsgBox2 and no real build. See issue #764.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")
'@Tag("unit")

Private Const CURRENT_DB As String = "C:\work\Scratch.mdb"
Private Const SOURCE_DB As String = "C:\repo\MyApp.mdb"
Private Const ALT_DB As String = "C:\work\MyApp.mdb"
Private Const SOURCE_SAME_NAME As String = "C:\repo\Scratch.mdb"


Public Sub TestNameConflictSkippedWhenNoDatabaseOpen()
    TestAssert Not ShouldCheckBuildNameConflict(vbNullString, SOURCE_DB, vbNullString, True), _
        "no current database skips the check"
End Sub


Public Sub TestNameConflictSkippedWhenPathsMatch()
    TestAssert Not ShouldCheckBuildNameConflict(SOURCE_DB, SOURCE_DB, vbNullString, True), _
        "matching paths skip the check on a regular full build"
End Sub


Public Sub TestNameConflictPromptedOnRegularFullBuild()
    TestAssert ShouldCheckBuildNameConflict(CURRENT_DB, SOURCE_DB, vbNullString, True), _
        "regular full build prompts when current and source paths differ"
End Sub


Public Sub TestNameConflictRefusedOnMerge()
    TestAssert ShouldCheckBuildNameConflict(CURRENT_DB, SOURCE_DB, vbNullString, False), _
        "merge refuses when current and source paths differ"
    TestAssert ShouldCheckBuildNameConflict(CURRENT_DB, SOURCE_DB, ALT_DB, False), _
        "merge still refuses even if an alternate path were passed"
End Sub


Public Sub TestNameConflictSkippedOnBuildAs()
    TestAssert Not ShouldCheckBuildNameConflict(CURRENT_DB, SOURCE_DB, ALT_DB, True), _
        "Build As skips the check when an alternate path is set"
End Sub


Public Sub TestNameConflictSkippedOnBuildAsSameBasenameDifferentFolder()
    TestAssert Not ShouldCheckBuildNameConflict(CURRENT_DB, SOURCE_SAME_NAME, ALT_DB, True), _
        "Build As skips even when the basename matches in a different folder"
End Sub


Public Sub TestNameConflictSkippedWhenBuildAsTargetsSourcePath()
    TestAssert Not ShouldCheckBuildNameConflict(CURRENT_DB, SOURCE_DB, SOURCE_DB, True), _
        "Build As skips even when the destination is the source-configured file"
End Sub
