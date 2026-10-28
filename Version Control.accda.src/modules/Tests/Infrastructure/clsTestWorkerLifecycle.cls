VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestWorkerLifecycle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestWorkerLifecycle
' Author    : Adam Waller
' Date      : 9/7/2026
' Purpose   : Queue-lifecycle tests for the worker file result channel, plus an
'           : end-to-end probe against the current database. Uses Friend seams so
'           : timeout and late-write cases do not log session errors.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.Infrastructure")
'@Tag("unit")

Private m_strTempFolder As String


Private Sub Class_Initialize()
    m_strTempFolder = GetTempFolder("vcs_worker_result")
End Sub


Private Sub Class_Terminate()
    If Len(m_strTempFolder) > 0 Then
        If FSO.FolderExists(m_strTempFolder) Then FSO.DeleteFolder StripSlash(m_strTempFolder), True
    End If
End Sub


Public Sub TestQueueFileResultCompletesAndDeletesFile()

    Dim cWorker As clsWorker
    Dim cJob As clsJob
    Dim strFile As String

    Set cWorker = New clsWorker
    strFile = WriteTokenFile("1")
    Set cJob = cWorker.QueueFileResultJob("lifecycle-complete", "CheckDatabaseAccessible", strFile)
    TestAssert Not cJob Is Nothing, "job was queued"
    cWorker.PollJobResults

    TestAssert cJob.Result = True, "1 completed as accessible"
    TestAssert Not cWorker.Queue.Exists(cJob.KeyHash), "queue entry removed"
    TestAssert Not FSO.FileExists(strFile), "result file deleted after completion"

End Sub


Public Sub TestQueueFileResultUnknownCompletes()

    Dim cWorker As clsWorker
    Dim cJob As clsJob
    Dim strFile As String

    Set cWorker = New clsWorker
    strFile = WriteTokenFile("U")
    Set cJob = cWorker.QueueFileResultJob("lifecycle-unknown", "CheckDatabaseAccessible", strFile)
    cWorker.PollJobResults

    TestAssert IsEmpty(cJob.Result), "U completes as Empty"
    TestAssert Not cWorker.Queue.Exists(cJob.KeyHash), "queue entry removed after unknown"

End Sub


Public Sub TestQueueFileResultIgnoresPartialContent()

    Dim cWorker As clsWorker
    Dim cJob As clsJob
    Dim strFile As String

    Set cWorker = New clsWorker
    strFile = WriteTokenFile("10")
    Set cJob = cWorker.QueueFileResultJob("lifecycle-partial", "CheckDatabaseAccessible", strFile)
    cWorker.PollJobResults

    TestAssert IsEmpty(cJob.Result), "partial content is not a verdict"
    TestAssert cWorker.Queue.Exists(cJob.KeyHash), "job stays queued for retry"
    TestAssert FSO.FileExists(strFile), "file kept until a real token arrives"

    WriteFile "0", strFile
    cWorker.PollJobResults
    TestAssert cJob.Result = False, "later 0 is accepted"
    TestAssert Not cWorker.Queue.Exists(cJob.KeyHash), "queue cleared after the real token"

End Sub


Public Sub TestDuplicateEnqueueLeavesFirstJob()

    Dim cWorker As clsWorker
    Dim cFirst As clsJob
    Dim cSecond As clsJob
    Dim strFile As String

    Set cWorker = New clsWorker
    strFile = WriteTokenFile(vbNullString)
    Set cFirst = cWorker.QueueFileResultJob("lifecycle-dup", "CheckDatabaseAccessible", strFile)
    Set cSecond = cWorker.QueueFileResultJob("lifecycle-dup", "CheckDatabaseAccessible", strFile)

    TestAssert Not cFirst Is Nothing, "first enqueue succeeds"
    TestAssert cSecond Is Nothing, "duplicate key is rejected without replacing the first"
    TestAssert cWorker.Queue.Count = 1, "only one job is queued"

End Sub


Public Sub TestExpireJobIgnoresLatePublication()

    Dim cWorker As clsWorker
    Dim cExpired As clsJob
    Dim cLater As clsJob
    Dim strExpired As String
    Dim strLater As String

    Set cWorker = New clsWorker
    strExpired = WriteTokenFile(vbNullString)
    Set cExpired = cWorker.QueueFileResultJob("lifecycle-late", "CheckDatabaseAccessible", strExpired)
    cWorker.ExpireJob cExpired

    TestAssert Not cWorker.Queue.Exists(cExpired.KeyHash), "expired job left the queue"
    TestAssert Not FSO.FileExists(strExpired), "reserved file deleted on expire"

    ' Simulate a late worker write to the old path after expiry.
    WriteFile "1", strExpired

    strLater = WriteTokenFile(vbNullString)
    Set cLater = cWorker.QueueFileResultJob("lifecycle-later", "CheckDatabaseAccessible", strLater)
    cWorker.PollJobResults

    TestAssert IsEmpty(cLater.Result), "later job is not completed by the stale file"
    TestAssert cWorker.Queue.Exists(cLater.KeyHash), "later job is still waiting on its own file"
    TestAssert cWorker.Queue.Count = 1, "expired job was not resurrected"

    cWorker.ExpireJob cLater
    If FSO.FileExists(strExpired) Then DeleteFile strExpired

End Sub


Public Sub TestReserveJobResultFileIsExclusive()

    Dim cWorker As clsWorker
    Dim strFirst As String
    Dim strSecond As String

    Set cWorker = New clsWorker
    strFirst = cWorker.ReserveJobResultFile("CheckDatabaseAccessible")
    strSecond = cWorker.ReserveJobResultFile("CheckDatabaseAccessible")

    TestAssert Len(strFirst) > 0, "first reservation produced a path"
    TestAssert Len(strSecond) > 0, "second reservation produced a path"
    TestAssert strFirst <> strSecond, "reservations do not share a name"
    TestAssert FSO.FileExists(strFirst), "first path was created exclusively"
    TestAssert FSO.FileExists(strSecond), "second path was created exclusively"

    DeleteFile strFirst
    DeleteFile strSecond

End Sub


Public Sub TestIsDatabaseAccessibleFileChannelCompletes()

    Dim cWorker As clsWorker
    Dim varResult As Variant

    If Not modInstall.UseWorkerScript Then
        Set cWorker = New clsWorker
        varResult = cWorker.IsDatabaseAccessible
        TestAssert IsEmpty(varResult), "disabled worker returns Empty without launching"
        Exit Sub
    End If

    Set cWorker = New clsWorker
    varResult = cWorker.IsDatabaseAccessible
    TestAssert Not IsEmpty(varResult), "probe returned a verdict through the file channel"
    TestAssert varResult = True Or varResult = False, "verdict is 1 or 0"
    TestAssert cWorker.Queue.Count = 0, "queue is empty after the probe"

End Sub


Private Function WriteTokenFile(strToken As String) As String
    WriteTokenFile = m_strTempFolder & PathSep & UniqueHashSuffix("worker-result") & ".result"
    ' WriteFile deletes the path when the text is empty; keep a non-verdict
    ' placeholder so PollJobResults can observe the file.
    If Len(strToken) = 0 Then
        WriteFile " ", WriteTokenFile
    Else
        WriteFile strToken, WriteTokenFile
    End If
End Function
