Attribute VB_Name = "modTestHeadlessBuild"
'---------------------------------------------------------------------------------------
' Module    : modTestHeadlessBuild
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Regression tests for the BuildHeadless / MergeHeadless preflight checks
'           : and JSON result shape.
'           :
'           : These deliberately only exercise the paths that refuse to start. A test
'           : that actually built or merged would rewrite the database it is running
'           : in, and there is no way back from that inside a test run.
'           :
'           : What matters to a pipeline is that a refusal still arrives as parseable
'           : JSON saying success=false with a reason, and that it does not leave an
'           : operation running -- a stuck operation makes every later VCS call in the
'           : session fail with "another operation is already running."
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Core")

' A path that holds no source files. Deliberately not GetTempFolder, which creates a
' folder that this module has no teardown to remove.
Private Const NO_SOURCE_FOLDER As String = "C:\vcs-no-such-source-folder\"


Public Sub TestBuildHeadlessRefusalIsWellFormed()

    Dim dResult As Dictionary

    Set dResult = ParseJson(VCS.BuildHeadless(NO_SOURCE_FOLDER))
    TestAssert Not CBool(dResult("success")), "folder without vcs-options.json fails"
    TestAssert Len(CStr(dResult("error"))) > 0, "failure carries a reason"
    TestAssert Operation.Status <> eosRunning, "refusal does not leave an operation running"

End Sub


Public Sub TestMergeHeadlessRefusalIsWellFormed()

    Dim dResult As Dictionary

    Set dResult = ParseJson(VCS.MergeHeadless(NO_SOURCE_FOLDER))
    TestAssert Not CBool(dResult("success")), "folder without vcs-options.json fails"
    TestAssert Len(CStr(dResult("error"))) > 0, "failure carries a reason"
    TestAssert Operation.Status <> eosRunning, "refusal does not leave an operation running"

End Sub


Public Sub TestHeadlessBuildRefusesToBuildTheAddInItself()

    Dim dResult As Dictionary

    ' Only meaningful when the suite is running inside the add-in. Against a user
    ' database the guard cannot trigger, and the folder check answers first.
    If StrComp(CodeProject.FullName, CurrentProject.FullName, vbTextCompare) <> 0 Then Exit Sub

    Set dResult = ParseJson(VCS.BuildHeadless(NO_SOURCE_FOLDER))
    TestAssert InStr(CStr(dResult("error")), "RebuildAddIn") > 0, _
        "refusal points at the supported way to rebuild the add-in"

End Sub


Public Sub TestHeadlessBuildRestoresInteractionMode()

    Dim intPrior As eInteractionMode

    ' A refused build must not leave the session silent, which would swallow every
    ' later prompt for someone who ran this from the Immediate Window.
    intPrior = Operation.InteractionMode
    VCS.BuildHeadless NO_SOURCE_FOLDER
    TestAssert Operation.InteractionMode = intPrior, "interaction mode restored"

End Sub
