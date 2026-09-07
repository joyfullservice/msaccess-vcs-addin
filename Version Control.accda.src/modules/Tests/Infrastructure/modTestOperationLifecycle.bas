Attribute VB_Name = "modTestOperationLifecycle"
'---------------------------------------------------------------------------------------
' Module    : modTestOperationLifecycle
' Author    : Adam Waller
' Date      : 8/20/2026
' Purpose   : Tests for root operation leases, detached continuations, pause scopes, and
'           : the interaction policy.
'           : These tests run *inside* a test-run root, so they must never call
'           : TryBeginRoot on the session operation: that either fails (a root is
'           : already active) or, worse, completes the harness's own root and tears
'           : down Log and TestRunner mid-run. Root behavior is therefore exercised on
'           : a private clsOperation instance, which performs no registry writes, MCP
'           : callbacks, VBE error-trapping changes, or singleton teardown.
'           : Run: ?VCS.RunTests("modTestOperationLifecycle")
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")


Private Const ModuleName As String = "modTestOperationLifecycle"


'---------------------------------------------------------------------------------------
' Procedure : NewOperation
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : A private operation instance, isolated from the session singleton.
'---------------------------------------------------------------------------------------
'
Private Function NewOperation(Optional blnUnattended As Boolean = False) As clsOperation
    Dim cOp As clsOperation
    Set cOp = New clsOperation
    cOp.ForceUnattended = blnUnattended
    Set NewOperation = cOp
End Function


Public Sub TestPrivateInstanceStartsReady()
    Dim cOp As clsOperation
    Set cOp = NewOperation
    TestAssert cOp.Status = eosReady, "private instance ignores restored root state"
    TestAssert Len(cOp.CurrentRootToken) = 0, "private instance holds no root token"
End Sub


Public Sub TestRootLeaseCompletesExactlyOnce()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotOther)
    TestAssert Not cRoot Is Nothing, "root lease acquired"
    TestAssert cOp.Status = eosRunning, "status running"
    cRoot.Complete eorSuccess
    TestAssert cOp.Status = eosReady, "status ready after complete"
    cRoot.Complete eorFailed
    TestAssert cOp.Result = eorSuccess, "second complete cannot change the result"
End Sub


Public Sub TestBeginHoldsRootUntilFinish()
    Dim cOp As clsOperation

    ' Begin discards no ownership: a lease released when Begin returns would abandon
    ' the root before any work starts, leaving Status eosReady for the whole
    ' operation and silently disabling every "is something running?" check.
    Set cOp = NewOperation(True)
    TestAssert cOp.Begin(eotExport), "synchronous root began"
    TestAssert cOp.Status = eosRunning, "root still running after Begin returns"
    TestAssert cOp.Result <> eorFailed, "root was not abandoned by its own lease"
    cOp.Finish eorSuccess
    TestAssert cOp.Status = eosReady, "Finish completed the root"
    TestAssert cOp.Result = eorSuccess, "Finish recorded the result"
End Sub


Public Sub TestSecondRootRefusedWhileFirstIsActive()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim cOther As clsRootOperationLease

    Set cOp = NewOperation(True)
    Set cRoot = cOp.TryBeginRoot(eotOther)
    TestAssert Not cRoot Is Nothing, "root lease acquired"
    Set cOther = cOp.TryBeginRoot(eotOther)
    TestAssert cOther Is Nothing, "second root refused"
    cRoot.Complete eorSuccess
    TestAssert cOp.Status = eosReady, "first lease still owns completion"
End Sub


Public Sub TestRootLeaseRejectsMismatchedComplete()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotOther)
    cOp.CompleteRootLease "not-a-real-token", eorFailed
    TestAssert cOp.Status = eosRunning, "bad token did not finish root"
    cRoot.Complete eorSuccess
    TestAssert cOp.Status = eosReady, "valid lease finished root"
End Sub


Public Sub TestAbandonedLeaseFinishesRoot()
    Dim cOp As clsOperation

    Set cOp = NewOperation
    AbandonLeaseInNestedScope cOp
    TestAssert cOp.Status = eosReady, "dropped lease released the root"
    TestAssert cOp.Result = eorFailed, "abandoned root records failure"
End Sub


Private Sub AbandonLeaseInNestedScope(cOp As clsOperation)
    ' The lease goes out of scope here without Complete, which is what Class_Terminate
    ' is the safety net for.
    Dim cRoot As clsRootOperationLease
    Set cRoot = cOp.TryBeginRoot(eotOther)
    TestAssert Not cRoot Is Nothing, "root lease acquired"
End Sub


Public Sub TestDetachResumeRequiresMatchingToken()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim cResumed As clsRootOperationLease
    Dim strToken As String

    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotMerge)
    strToken = cRoot.DetachForContinuation()
    TestAssert cOp.Status = eosStaged, "detached root is staged"
    TestAssert Len(strToken) > 0, "continuation token issued"

    Set cResumed = cOp.ResumeRoot("not-the-token")
    TestAssert cResumed Is Nothing, "mismatched token refused"
    TestAssert cOp.Status = eosStaged, "refused resume left the root staged"

    Set cResumed = cOp.ResumeRoot(strToken)
    TestAssert Not cResumed Is Nothing, "matching token resumes root"
    TestAssert cOp.Status = eosRunning, "resumed root is running"
    cResumed.Complete eorSuccess
    TestAssert cOp.Status = eosReady, "resume complete leaves ready"
End Sub


Public Sub TestResumeRefusedAfterCompletion()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim strToken As String

    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotBuild)
    strToken = cRoot.Token
    cRoot.Complete eorSuccess
    TestAssert cOp.ResumeRoot(strToken) Is Nothing, "completed root cannot be resumed"
End Sub


Public Sub TestSecondDetachMustTransferLeaseNotComplete()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim cResumed As clsRootOperationLease
    Dim cResumed2 As clsRootOperationLease
    Dim strToken As String

    ' Completing the resumed lease after a second detach finishes the root.
    ' That is the hang: in-place merge prep detaches again for MergeReset, and
    ' RunBuildFromContinuation used to Complete when Build returned.
    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotMerge)
    strToken = cRoot.DetachForContinuation()
    Set cResumed = cOp.ResumeRoot(strToken)
    TestAssert Not cResumed Is Nothing, "first resume succeeded"
    cOp.DetachRootLease strToken
    TestAssert cOp.Status = eosStaged, "second detach leaves root staged"
    cResumed.Complete eorSuccess
    TestAssert cOp.Status = eosReady, "Complete on the resumed lease finishes the root"
    TestAssert cOp.ResumeRoot(strToken) Is Nothing, "completed root cannot be resumed"

    ' Transferring the lease instead leaves the original token resumable.
    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotMerge)
    strToken = cRoot.DetachForContinuation()
    Set cResumed = cOp.ResumeRoot(strToken)
    TestAssert Not cResumed Is Nothing, "second-path first resume succeeded"
    cOp.DetachRootLease strToken
    cResumed.DetachForContinuation
    TestAssert cOp.Status = eosStaged, "transferred lease leaves root staged"
    TestAssert Not cResumed.IsValid, "original lease is no longer valid"
    Set cResumed2 = cOp.ResumeRoot(strToken)
    TestAssert Not cResumed2 Is Nothing, "second continuation can resume"
    TestAssert cOp.Status = eosRunning, "resumed after transfer is running"
    cResumed2.Complete eorSuccess
    TestAssert cOp.Status = eosReady, "transferred continuation still completes"
End Sub


Public Sub TestPauseScopeDoesNotFinishRoot()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim cPause As clsOperationPause

    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotOther)
    Set cPause = cOp.TryPause()
    TestAssert Not cPause Is Nothing, "pause scope acquired"
    TestAssert cOp.Status = eosStaged, "paused root is suspended, not canceled"
    TestAssert cRoot.IsValid, "lease survives the pause"
    cPause.ResumePause
    TestAssert cOp.Status = eosRunning, "root running again after the pause"
    cRoot.Complete eorSuccess
    TestAssert cOp.Status = eosReady, "lease still completes the root"
End Sub


Public Sub TestNestedPausesResumeOnlyAtTheOutermost()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim cOuter As clsOperationPause
    Dim cInner As clsOperationPause

    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotOther)
    Set cOuter = cOp.TryPause()
    Set cInner = cOp.TryPause()
    cInner.ResumePause
    TestAssert cOp.Status = eosStaged, "inner resume leaves the outer pause in effect"
    cOuter.ResumePause
    TestAssert cOp.Status = eosRunning, "outer resume restores the root"
    cRoot.Complete eorSuccess
End Sub


Public Sub TestPauseRefusedWithoutRunningRoot()
    Dim cOp As clsOperation
    Set cOp = NewOperation
    TestAssert cOp.TryPause() Is Nothing, "no pause without a running root"
End Sub


Public Sub TestRootFinishResetsInteractionState()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    Set cOp = NewOperation(True)
    cOp.InteractionMode = eimSilent
    Set cRoot = cOp.TryBeginRoot(eotOther)
    TestAssert Not cOp.Attended, "ForceUnattended applied at root creation"
    cRoot.Complete eorSuccess
    TestAssert cOp.InteractionMode = eimNormal, "interaction mode reset after finish"
    TestAssert Not cOp.ForceUnattended, "ForceUnattended cleared after finish"
End Sub


Public Sub TestSilentModeCannotBeRelaxedDuringRoot()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    Set cOp = NewOperation(True)
    Set cRoot = cOp.TryBeginRoot(eotTestRun)
    cOp.InteractionMode = eimSilent
    cOp.InteractionMode = eimNormal
    TestAssert cOp.InteractionMode = eimSilent, "relaxing is ignored while a root is active"
    cRoot.Complete eorSuccess
    cOp.InteractionMode = eimNormal
    TestAssert cOp.InteractionMode = eimNormal, "relaxing is allowed between roots"
End Sub


Public Sub TestAttendedIsImmutableForRootLifetime()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    Set cOp = NewOperation
    Set cRoot = cOp.TryBeginRoot(eotOther)
    TestAssert cOp.Attended, "interactive root is attended"
    cOp.ForceUnattended = True
    TestAssert cOp.Attended, "late ForceUnattended does not change the live root"
    cRoot.Complete eorSuccess
End Sub


Public Sub TestRunningRootExpiresAfterHeartbeatTimeout()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease

    ' A root that stops pulsing is how a crashed or abandoned operation is recognized,
    ' so the timeout has to still fire for a root that is nominally running.
    Set cOp = NewOperation(True)
    Set cRoot = cOp.TryBeginRoot(eotExport)
    cOp.Heartbeat = DateAdd("n", -30, Now)
    TestAssert cOp.Status = eosReady, "a root that stopped pulsing times out"
    TestAssert cOp.Result = eorTimeout, "the timeout is recorded as the result"
    cRoot.Complete eorTimeout
End Sub


Public Sub TestPausedRootIgnoresHeartbeatTimeout()
    Dim cOp As clsOperation
    Dim cRoot As clsRootOperationLease
    Dim cPause As clsOperationPause

    ' A user hook such as AfterExport can legitimately run past the timeout, and nothing
    ' can pulse while foreign code holds the stack. Expiring the root there would strand
    ' it, since EndPauseScope only restores a root it still finds staged.
    Set cOp = NewOperation(True)
    Set cRoot = cOp.TryBeginRoot(eotExport)
    Set cPause = cOp.TryPause()
    cOp.Heartbeat = DateAdd("n", -30, Now)
    TestAssert cOp.Status = eosStaged, "a long-running hook does not expire the paused root"
    cPause.ResumePause
    TestAssert cOp.Status = eosRunning, "resuming restores the root and refreshes the heartbeat"
    cRoot.Complete eorSuccess
    TestAssert cOp.Result = eorSuccess, "the root still completes normally"
End Sub


Public Sub TestObjectScanProgressPulsesHeartbeat()
    Dim sngStart As Single
    Dim datBefore As Date

    ' Change detection pulses once per category; this progress hook is the only call
    ' every scan loop makes per component. sngStart of 0 suppresses the console
    ' breadcrumb, so this touches nothing but the heartbeat of the run hosting it.
    Operation.Heartbeat = DateAdd("s", -2, Now)
    datBefore = Operation.Heartbeat
    Log.IncrementObjectScanProgress sngStart, 0, "Queries"
    TestAssert Operation.Heartbeat > datBefore, "scanning a component pulses the heartbeat"
End Sub


Public Sub TestHeadlessBuildRefusalLeavesSessionRootIntact()
    Dim dResult As Dictionary
    Dim strTokenBefore As String

    ' A refused headless build must neither prompt nor disturb the root that is running
    ' this very test.
    strTokenBefore = Operation.CurrentRootToken
    Set dResult = ParseJson(VCS.BuildHeadless("C:\vcs-no-such-source-folder\"))
    TestAssert Not CBool(dResult("success")), "invalid folder fails"
    TestAssert Operation.CurrentRootToken = strTokenBefore, "session root untouched"
End Sub
