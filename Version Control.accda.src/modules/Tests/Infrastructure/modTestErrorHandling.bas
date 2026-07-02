Attribute VB_Name = "modTestErrorHandling"
'---------------------------------------------------------------------------------------
' Module    : modTestErrorHandling
' Author    : Adam Waller
' Date      : 5/12/2026
' Purpose   : Tests for modErrorHandling: Catch, CatchAny, LogUnhandledErrors.
'           : Migrated from TestCatch in modTestSuite.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")

Private Const ModuleName As String = "modTestErrorHandling"


Public Sub TestCatch()
    Dim blnOrigBreak As Boolean
    Dim eimPriorMode As eInteractionMode
    blnOrigBreak = Options.BreakOnError
    eimPriorMode = Operation.InteractionMode
    Options.BreakOnError = False
    Operation.InteractionMode = eimSilent

    Log.Add "  [TestCatch] Deliberately raising errors to exercise error-handling functions. " & _
        "Any logged errors from this test are expected and safe to ignore."

    On Error Resume Next
    Err.Raise 24601, "Pre Log Test"

    ' LogUnhandledErrors should capture the error without crashing
    LogUnhandledErrors ModuleName & ".TestCatch (expected test error)"
    On Error Resume Next

    ' Raise another error and verify CatchAny handles it
    Err.Raise 24602, "Post Log Test"
    CatchAny eelError, "Expected test error - verifying CatchAny handles eelError", _
        ModuleName & ".TestCatch"

    ' If we got here, Catch/CatchAny didn't crash
    TestAssert True, "Catch and CatchAny completed without crash"

    Operation.InteractionMode = eimPriorMode
    Options.BreakOnError = blnOrigBreak
End Sub


Public Sub TestCatch_SpecificError()
    On Error Resume Next
    Err.Raise 13

    TestAssert Catch(13), "catches type mismatch (13)"
    TestAssert Not Catch(13), "error cleared after first Catch"
End Sub


Public Sub TestCatch_NoError()
    On Error Resume Next
    ' No error raised
    TestAssert Not Catch(13), "returns False when no error"
End Sub
