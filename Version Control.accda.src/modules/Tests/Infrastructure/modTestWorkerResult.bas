Attribute VB_Name = "modTestWorkerResult"
'---------------------------------------------------------------------------------------
' Module    : modTestWorkerResult
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 9/1/2026
' Purpose   : Unit tests for the worker's file result channel: what counts as an answer
'           : from the accessibility probe and what does not. The distinction matters
'           : because a read can fail transiently, and answering "not accessible" to
'           : that would be a false negative that is silent and intermittent.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Infrastructure")
'@Tag("unit")


Public Sub TestWorkerResultAcceptsOnlyTheTwoWrittenValues()

    Dim cWorker As clsWorker
    Dim blnAccessible As Boolean
    Dim blnIsVerdict As Boolean

    Set cWorker = New clsWorker

    blnIsVerdict = cWorker.JobResultFromContent("1", blnAccessible)
    TestAssert blnIsVerdict, "1 is a verdict"
    TestAssert blnAccessible, "1 means the database is accessible"

    blnIsVerdict = cWorker.JobResultFromContent("0", blnAccessible)
    TestAssert blnIsVerdict, "0 is a verdict"
    TestAssert Not blnAccessible, "0 means the database is not accessible"

    ' The worker writes the value on its own line, so a trailing newline is ordinary.
    blnIsVerdict = cWorker.JobResultFromContent("1" & vbCrLf, blnAccessible)
    TestAssert blnIsVerdict, "a trailing newline does not spoil the verdict"

    blnIsVerdict = cWorker.JobResultFromContent("  0  ", blnAccessible)
    TestAssert blnIsVerdict, "surrounding whitespace is trimmed"
    TestAssert Not blnAccessible, "and the trimmed value is still the one read"

End Sub


Public Sub TestWorkerResultTreatsAnythingElseAsNoAnswerYet()

    Dim cWorker As clsWorker
    Dim blnAccessible As Boolean

    Set cWorker = New clsWorker

    ' This value must survive every call below: the caller retries, and a non-answer
    ' that overwrote it would turn a slow read into a verdict nobody wrote.
    blnAccessible = True

    TestAssert Not cWorker.JobResultFromContent("", blnAccessible), _
        "an empty file is not an answer"
    TestAssert Not cWorker.JobResultFromContent("   ", blnAccessible), _
        "a blank line is not an answer"
    TestAssert Not cWorker.JobResultFromContent("10", blnAccessible), _
        "10 is not one of the two values the worker writes"
    TestAssert Not cWorker.JobResultFromContent("true", blnAccessible), _
        "the worker never writes words"
    TestAssert Not cWorker.JobResultFromContent(vbCrLf & "1", blnAccessible), _
        "the value is read from the first line only"

    TestAssert blnAccessible, "a non-answer leaves the caller value untouched"

End Sub

