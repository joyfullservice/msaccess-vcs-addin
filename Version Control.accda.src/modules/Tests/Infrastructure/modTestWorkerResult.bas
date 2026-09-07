Attribute VB_Name = "modTestWorkerResult"
'---------------------------------------------------------------------------------------
' Module    : modTestWorkerResult
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 9/1/2026
' Purpose   : Unit tests for the worker's file result channel: which tokens count as
'           : a terminal answer (1, 0, U) and which leave the caller untouched.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")


Public Sub TestWorkerResultAcceptsWrittenTokens()

    Dim cWorker As clsWorker
    Dim varResult As Variant
    Dim blnIsVerdict As Boolean

    Set cWorker = New clsWorker

    blnIsVerdict = cWorker.JobResultFromContent("1", varResult)
    TestAssert blnIsVerdict, "1 is a verdict"
    TestAssert varResult = True, "1 means the database is accessible"

    blnIsVerdict = cWorker.JobResultFromContent("0", varResult)
    TestAssert blnIsVerdict, "0 is a verdict"
    TestAssert varResult = False, "0 means the database is not accessible"

    ' The worker writes the value on its own line, so a trailing newline is ordinary.
    blnIsVerdict = cWorker.JobResultFromContent("1" & vbCrLf, varResult)
    TestAssert blnIsVerdict, "a trailing newline does not spoil the verdict"
    TestAssert varResult = True, "newline still reads as accessible"

    blnIsVerdict = cWorker.JobResultFromContent("  0  ", varResult)
    TestAssert blnIsVerdict, "surrounding whitespace is trimmed"
    TestAssert varResult = False, "and the trimmed value is still the one read"

    blnIsVerdict = cWorker.JobResultFromContent("U", varResult)
    TestAssert blnIsVerdict, "U is a published unknown"
    TestAssert IsEmpty(varResult), "U leaves the transport result Empty"

    blnIsVerdict = cWorker.JobResultFromContent("U" & vbCrLf, varResult)
    TestAssert blnIsVerdict, "U with a trailing newline is still unknown"
    TestAssert IsEmpty(varResult), "unknown survives the newline"

End Sub


Public Sub TestWorkerResultTreatsAnythingElseAsNoAnswerYet()

    Dim cWorker As clsWorker
    Dim varResult As Variant

    Set cWorker = New clsWorker

    ' This value must survive every call below: the caller retries, and a non-answer
    ' that overwrote it would turn a slow read into a verdict nobody wrote.
    varResult = True

    TestAssert Not cWorker.JobResultFromContent("", varResult), _
        "an empty file is not an answer"
    TestAssert Not cWorker.JobResultFromContent("   ", varResult), _
        "a blank line is not an answer"
    TestAssert Not cWorker.JobResultFromContent("10", varResult), _
        "10 is not one of the tokens the worker writes"
    TestAssert Not cWorker.JobResultFromContent("true", varResult), _
        "the worker never writes words"
    TestAssert Not cWorker.JobResultFromContent("unknown", varResult), _
        "the word unknown is not the U token"
    TestAssert Not cWorker.JobResultFromContent(vbCrLf & "1", varResult), _
        "the value is read from the first line only"

    TestAssert varResult = True, "a non-answer leaves the caller value untouched"

End Sub


Public Sub TestAccessibilityFromProbeResultCollapsesUnknown()

    Dim cWorker As clsWorker

    Set cWorker = New clsWorker

    TestAssert Not cWorker.AccessibilityFromProbeResult(Empty), _
        "unknown is not accessible"
    TestAssert cWorker.AccessibilityFromProbeResult(True), _
        "True stays accessible"
    TestAssert Not cWorker.AccessibilityFromProbeResult(False), _
        "False stays inaccessible"

    ' All three build callers go through DatabaseAccessibleToOtherClients, which
    ' uses this collapse. Unknown therefore cannot reach the in-place merge path,
    ' which requires a positive result before setting m_blnVerifiedAccessible.

End Sub
