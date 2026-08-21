Attribute VB_Name = "modTestPromptSuppression"
'---------------------------------------------------------------------------------------
' Module    : modTestPromptSuppression
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Tests that a harness-driven run cannot be blocked by a prompt.
'           : These run in the project being tested, which is not the project driving the
'           : run: the add-in exports its own source, so testing it loads a second copy of
'           : this project, and the operation that owns the run belongs to the other copy.
'           : Nothing here can see that operation, which is exactly the situation a user
'           : database is in. So the contract asserted here is the one every hosted
'           : project gets: modTestAssert.TestRunActive is true, and PromptWouldDisplay is
'           : false for every kind of prompt.
'           : MsgBox2 is only called where suppression is already established, and never
'           : with the user-gesture flag.
'           : Run: ?VCS.RunTests("modTestPromptSuppression")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")


Private Const ModuleName As String = "modTestPromptSuppression"


Public Sub TestRunAnnouncesItselfToTheHostedProject()
    TestAssert TestRunActive, "the run driving this project announced itself"
End Sub


Public Sub TestRunSuppressesOrdinaryPrompts()
    TestAssert Not PromptWouldDisplay(False), "ordinary prompt suppressed during a test run"
End Sub


Public Sub TestRunSuppressesGesturePromptsToo()
    ' A gesture prompt outlives silent mode when a person is driving, but during a test
    ' run the runner is the caller: there is no gesture, and a dialog stalls the suite.
    TestAssert Not PromptWouldDisplay(True), "gesture prompt suppressed during a test run"
End Sub


Public Sub TestSuppressionOutranksTheLocalOperation()
    Dim eimPrior As eInteractionMode

    ' No root is active in this project, so relaxing the local operation is allowed here.
    ' Suppression must not depend on it.
    eimPrior = Operation.InteractionMode
    Operation.InteractionMode = eimNormal
    TestAssert Not PromptWouldDisplay(False), "prompts stay suppressed under a relaxed local operation"
    Operation.InteractionMode = eimPrior
End Sub


Public Sub TestSilentMsgBox2ReturnsDefaultResult()
    Dim intResult As VbMsgBoxResult

    ' The only live MsgBox2 call in the suite, and it runs only once suppression is
    ' confirmed. If the precondition fails the call is skipped rather than attempted:
    ' a real dialog here stalls the whole run until someone clicks it.
    If PromptWouldDisplay(False) Then
        TestAssert False, "precondition failed: prompts are not suppressed, MsgBox2 call skipped"
        Exit Sub
    End If

    TestAssert True, "precondition: prompts are suppressed"
    intResult = MsgBox2("Suppressed prompt", "line1", , vbYesNo, , vbNo)
    TestAssert intResult = vbNo, "suppressed MsgBox2 returns the unattended default"
End Sub
