VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsWorker
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : This module has two primary purposes.
'           :
'           : 1. To run commands in the `current database` that would otherwise risk
'           : running in the add-in database. (Deleting a database object, for example,
'           : will default to the CodeDB before CurrentDB.)
'           :
'           : 2. Opens up the opportunity for parallel execution when VBScript can run
'           : operations asynchrously. (For example, scanning through a large number of
'           : source files can be slow, but this gives us the ability to scan multiple
'           : folders at the same time which should improve performance on SSD drives.)
'           :
'           : Code below the ' *** BEGIN WORKER SCRIPT *** marker is extracted verbatim
'           : into Worker.vbs and runs as VBScript, not VBA. See that banner for rules.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Integration")


' Worker Job queue of jobs in progress
Public Queue As Dictionary

' Last result returned by a worker callback
Private m_varLastResult As Variant

' Result file and queue key of the job reporting through the file channel, if any. One
' field each for the same reason m_varLastResult is one field: only one job at a time can
' be waited on for a value, and the duplicate-job guard in CallWorker keeps it that way.
Private m_strResultFile As String
Private m_strResultKey As String

' Generic placeholder objects so we can compile the VBScript in the VBA IDE
Private WScript As Object


'---------------------------------------------------------------------------------------
' Procedure : Run_UninstallAddin
' Author    : Adam Waller
' Date      : 5/22/2023
' Purpose   : Remove the actual add-in file and lock file during uninstall.
'---------------------------------------------------------------------------------------
'
Public Function Run_UninstallAddin()
    CallWorker PROJECT_NAME, 10, "Uninstall", """" & GetInstalledAddInFileName & """"
End Function


'---------------------------------------------------------------------------------------
' Procedure : Run_BuildAndInstall
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Rebuild the VCS add-in from source files and install the result.
'           : Runs disconnected: waits for the current Access instance to exit, opens a
'           : new one to perform the build, then launches the install via /cmd INSTALL.
'           : blnUnattended suppresses worker MsgBox prompts and uses /cmd INSTALL SILENT.
'           :
'           : The Access executable path and this process ID are passed as arguments
'           : rather than read back off a COM reference, because this is the one job
'           : that must still run when the worker cannot attach to its caller. See the
'           : attach in Main.
'---------------------------------------------------------------------------------------
'
Public Function Run_BuildAndInstall(strSourcePath As String, _
    Optional blnUnattended As Boolean = False, _
    Optional strStatusFile As String, _
    Optional strPhaseStarted As String, _
    Optional strCallbackUrl As String, _
    Optional strOperationId As String)

    CallWorker PROJECT_NAME, 300, "BuildAndInstall", _
        """" & strSourcePath & """", _
        IIf(blnUnattended, "1", "0"), _
        """" & strStatusFile & """", _
        """" & FSO.BuildPath(SysCmd(acSysCmdAccessDir), "MSACCESS.EXE") & """", _
        modInstall.GetThisProcessId, _
        """" & strPhaseStarted & """", _
        """" & strCallbackUrl & """", _
        """" & strOperationId & """"

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsDatabaseAccessible
' Author    : Adam Waller
' Date      : 4/2/2026
' Purpose   : Check if the database file is accessible to external clients by running
'           : an out-of-process DAO connection test via the worker script. This detects
'           : the JET/ACE engine-level lock state (error 3734) that an in-process check
'           : cannot detect.
'---------------------------------------------------------------------------------------
'
Public Function IsDatabaseAccessible() As Boolean

    ' This action reports through a file instead of the callback. It is the only one that
    ' does not need the Access instance to do its work -- it builds its own DAO engine --
    ' so attaching would be a dependency taken out purely in order to answer, and that is
    ' the attach that cannot succeed when the caller's current database is the add-in
    ' itself. See the 2026-08-12 entry in DECISIONS.md for the measured 432.
    m_varLastResult = Empty
    m_strResultKey = JobKeyHash("DbAccessCheck", "CheckDatabaseAccessible")
    m_strResultFile = JobResultFile("CheckDatabaseAccessible")
    CallWorker "DbAccessCheck", 10, "CheckDatabaseAccessible", """" & m_strResultFile & """"
    Me.WaitForQueue 10, "CheckDatabaseAccessible", 0.1

    ' Three outcomes, not two. No file at all means the worker never answered, which is
    ' not the same as an answer of "not accessible" even though the caller takes the same
    ' conservative action either way. Stay quiet when the helper script is switched off
    ' (#727): there an empty result is the documented behaviour, not a fault.
    If IsEmpty(m_varLastResult) Then
        If modInstall.UseWorkerScript Then
            Log.Error eelWarning, T("No result from the worker. Assuming the database is not accessible to other clients."), _
                ModuleName(Me) & ".IsDatabaseAccessible"
        End If
        IsDatabaseAccessible = False
    Else
        IsDatabaseAccessible = CBool(m_varLastResult)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : Run_SaveVbaProject
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Save the current database's VBA project from outside the VBA stack, which is
'           : the only way the VBE Save command actually saves it. In process the same
'           : command reports success and leaves the project dirty. Returns the resulting
'           : Saved state. See modVbeUtility.SaveCurrentVBProject for what else was tried,
'           : including resetting the project in this same job (reverted -- it broke export).
'---------------------------------------------------------------------------------------
'
Public Function Run_SaveVbaProject() As Boolean
    m_varLastResult = Empty
    CallWorker "VbaProjectSave", 30, "SaveVbaProject"
    Me.WaitForQueue 30, "SaveVbaProject", 0.1
    Run_SaveVbaProject = CBool(m_varLastResult)
End Function


'---------------------------------------------------------------------------------------
' Procedure : DeleteDatabaseObject
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Delete a database object from the current database. (The reason we are
'           : not just using the native command is because when it is initiated from
'           : the add-in, it will delete any object with the same name in the add-in
'           : before deleting the one in the current database. This will cause the
'           : add-in file to be currupted, and the operation will fail.)
'---------------------------------------------------------------------------------------
'
Public Function WIP_DeleteDatabaseObject(intType As AcObjectType, strName As String, Optional blnWaitForQueue As Boolean = True)

    ' Don't use the following method from the add-in! Instead, we need to call
    ' it from an external process so it deletes from the current database first.
    'DoCmd.DeleteObject acModule, "basModule"
    CallWorker strName, 5, "DeleteDatabaseObject", intType, strName
    If blnWaitForQueue Then Me.WaitForQueue 5, "DeleteDatabaseObject", 0.25

End Function


'---------------------------------------------------------------------------------------
' Procedure : RenameDatabaseObject
' Author    : Adam Waller
' Date      : 3/3/2023
' Purpose   : Rename the specified database object
'---------------------------------------------------------------------------------------
'
Public Function WIP_RenameDatabaseObject(strNewName As String, intType As AcObjectType, _
    strOldName As String, Optional blnWaitForQueue As Boolean = True)
    CallWorker strNewName, 3, "RenameDatabaseObject", strNewName, intType, strOldName
    If blnWaitForQueue Then Me.WaitForQueue 3, "RenameDatabaseObject", 0.1
End Function


'---------------------------------------------------------------------------------------
' Procedure : WaitForQueue
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Wait for the job queue to finish processing, checking at each specified
'           : interval. (Don't make the poll interval too high, or the workers may not
'           : be able to complete their callbacks.)
'---------------------------------------------------------------------------------------
'
Public Sub WaitForQueue(dblTimeout As Double, Optional strAction As String = "All", Optional sngPollSeconds As Single = 0.5)

    Dim cJob As clsJob
    Dim varKey As Variant
    Dim blnFound As Boolean
    Dim dblStart As Double
    Dim dblPollStart As Double
    Dim colRemove As Collection
    Dim varJob As Variant

    ' No need to wait if there is nothing in the queue
    If Me.Queue.Count = 0 Then Exit Sub

    ' Get start time so we can monitor the timeout
    dblStart = Perf.MicroTimer
    Set colRemove = New Collection

    ' Monitor performance of waiting for the job queue to clear
    Perf.OperationStart "Wait for Job Queue"

    ' Loop while periodically checking the queue
    Do
        ' Pause for the poll interval
        dblPollStart = Perf.MicroTimer
        Do While dblPollStart + sngPollSeconds > Perf.MicroTimer
            ' A tight DoEvents loop is required here. The worker
            ' VBScript makes many individual COM calls back into
            ' Access (property access, VBE iteration, etc.), each
            ' needing message pump dispatch. Sleep would add ~100ms
            ' latency per round-trip, compounding to seconds of delay.
            DoEvents
        Loop
        DoEvents

        ' Pick up a result delivered through the file channel. An action that does not
        ' attach has no callback to make, so nothing arrives here on its own.
        PollJobResults

        ' We can exit early if there is nothing in the queue
        If Me.Queue.Count = 0 Then Exit Do

        ' Reset flag for this iteration
        blnFound = False

        ' Loop through queue of jobs
        For Each varKey In Me.Queue.Keys
            Set cJob = Me.Queue(varKey)

            ' Check for job timeout (all jobs)
            If Perf.MicroTimer > cJob.Timeout Then
                ' Add to list of items to remove
                colRemove.Add CStr(varKey)
            Else
                ' Look for matching actions
                If cJob.Action = strAction Then
                    blnFound = True
                    Exit For
                End If
            End If
        Next varKey
        Set cJob = Nothing

        ' Log and remove any individual jobs that have timed out
        For Each varJob In colRemove
            Set cJob = Me.Queue(varJob)
            With cJob
                Log.Error eelError, "Job " & .KeyHash & " (" & .Action & ") timed out after " & _
                    Round(Perf.MicroTimer - .Start, 0) & " seconds.", ModuleName(Me) & ".WaitForQueue"
            End With
            Set cJob = Nothing
            If Me.Queue.Exists(varJob) Then Me.Queue.Remove varJob
        Next varJob

        ' If nothing found of a specified action, we can exit the loop
        If (strAction <> "All") And (Not blnFound) Then Exit Do

        ' Check the overall timeout before we start another iteration
        If dblStart + dblTimeout < Perf.MicroTimer Then
            Log.Error eelError, "Timed out waiting for job queue", ModuleName(Me) & ".WaitForQueue"
            Log.Add " (" & Me.Queue.Count & " jobs still in the queue.)", False
            Exit Do
        End If
    Loop
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : VerifyWorker
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Verify that the work script is installed and unchanged.
'---------------------------------------------------------------------------------------
'
Private Function VerifyWorker() As Boolean

    Static blnVerified As Boolean

    Dim strContent As String
    Dim strFilePath As String

    ' Only need to verify this once per session.
    If blnVerified Then
        VerifyWorker = True
        Exit Function
    End If

    ' Compare hash of file with hash of content
    strContent = GetWorkerScriptContent
    strFilePath = WorkerFilePath
    If FSO.FileExists(strFilePath) Then
        If GetFileHash(WorkerFilePath) = GetStringHash(strContent, True) Then
            ' Existing file matches template.
            blnVerified = True
        End If
    End If

    ' Update the file if needed.
    If Not blnVerified Then
        WriteFile strContent, strFilePath, "Windows-1252"
        blnVerified = True
    End If
    VerifyWorker = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetWorkerScriptContent
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Return the content of the script, ready to save to a `Worker.vbs` file.
'---------------------------------------------------------------------------------------
'
Private Function GetWorkerScriptContent() As String

    Dim strContent As String
    Dim varLines As Variant
    Dim strLine As String
    Dim blnSave As Boolean
    Dim lngLine As Long

    ' Get VBA code from this class
    strContent = GetCodeVBProject.VBComponents(TypeName(Me)).CodeModule.Lines(1, 99999)

    ' Split into lines
    varLines = Split(strContent, vbCrLf)

    ' Look for the marker where the placeholder starts, and keep everything after that.
    With New clsConcat
        For lngLine = 0 To UBound(varLines) - 1
            strLine = varLines(lngLine)
            If blnSave Then
                .Add strLine, vbCrLf
            ElseIf strLine = "' *** BEGIN WORKER SCRIPT ***" Then
                blnSave = True
                ' Add line to top of VBScript to launch the main sub
                .Add "Main ' Launch main subroutine", vbCrLf, vbCrLf
            End If
        Next lngLine
        ' Return script content
        GetWorkerScriptContent = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : WorkerFilePath
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : File path to the installed worker script.
'---------------------------------------------------------------------------------------
'
Private Function WorkerFilePath() As String
    WorkerFilePath = FSO.GetParentFolderName(modInstall.GetAddInFileName) _
        & PathSep & "Worker.vbs"
End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveWorkerScript
' Author    : Adam Waller
' Date      : 8/7/2026
' Purpose   : Delete the extracted worker script, if it is present. Used when the user
'           : turns the helper script off, and during uninstall. For a user whose
'           : endpoint protection objects to the script (#727), leaving the file sitting
'           : in the add-in folder is part of the problem, so disabling the feature
'           : removes the artifact rather than merely stopping its use.
'           : Lives here so the file path is defined in exactly one place.
'---------------------------------------------------------------------------------------
'
Public Sub RemoveWorkerScript()
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    ' DeleteFile is a no-op when the file is already gone.
    DeleteFile WorkerFilePath, True
    CatchAny eelWarning, T("Unable to remove worker script: {0}", var0:=WorkerFilePath), _
        ModuleName(Me) & ".RemoveWorkerScript"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CallWorker
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Call the worker script to run a specified operation.
'---------------------------------------------------------------------------------------
'
Private Function CallWorker(strKey As String, dblTimeoutSeconds As Double, _
    strAction As String, ParamArray varParams()) As Long

    Dim intParam As Integer
    Dim strHash As String
    Dim lngHandle As Long
    Dim cJob As clsJob

    ' The user can turn the helper script off when their environment does not allow
    ' Access to launch an extracted script (#727). Every job becomes a no-op, which
    ' means a caller that forgets to branch degrades instead of launching wscript.
    ' Nothing is added to the queue, so a caller that waits returns straight away and
    ' reads an empty result -- see the fallbacks in modBuild and modVbeUtility.
    ' This is also the seam for any future out-of-process backend: keeping the launch
    ' in one place means a replacement would not touch a single call site.
    If Not modInstall.UseWorkerScript Then
        Log.Add "Worker script is disabled. Skipping job: " & strAction, False
        Exit Function
    End If

    ' Make sure this key/action is not already running
    strHash = JobKeyHash(strKey, strAction)
    If Me.Queue.Exists(strHash) Then
        Log.Error eelError, "Duplicate job in queue: " & strAction & "." & strKey, _
            ModuleName(Me) & ".CallWorker"
        Exit Function
    End If

    ' Verify that the work script is installed and current
    VerifyWorker

    ' Build command with parameters
    With New clsConcat

        ' Build base command to launch worker
        .Add "wscript """, WorkerFilePath, """"

        ' Add database file as first parameter
        .Add " """, CurrentProject.FullName, """"

        ' Add unique key as second parameter
        .Add " ", strHash

        ' Add specified action
        .Add " ", strAction

        ' Add any parameters
        For intParam = 0 To UBound(varParams)
            .Add " ", CStr(varParams(intParam))
        Next intParam

        ' Add to job queue (before launching, just in case it runs very fast)
        Set cJob = New clsJob
        With cJob
            .KeyHash = strHash
            .Action = strAction
            .Start = Perf.MicroTimer
            .Timeout = .Start + dblTimeoutSeconds
        End With
        Me.Queue.Add strHash, cJob

        ' Use the Shell command to launch the worker as an external asynchrous
        ' process, and return the process ID to the caller.
        lngHandle = Shell(.GetStr, vbHide)

    End With

    ' Return process handle
    CallWorker = lngHandle

End Function


'---------------------------------------------------------------------------------------
' Procedure : JobKeyHash
' Author    : Adam Waller
' Date      : 8/23/2026
' Purpose   : The queue key for a job. Deterministic on purpose: it is what makes the
'           : duplicate check work, and what the worker echoes back in a callback.
'---------------------------------------------------------------------------------------
'
Private Function JobKeyHash(strKey As String, strAction As String) As String
    JobKeyHash = GetStringHash(strKey & strAction)
End Function


'---------------------------------------------------------------------------------------
' Procedure : JobResultFile
' Author    : Adam Waller
' Date      : 8/23/2026
' Purpose   : Path the worker writes a job's return value to, handed to it as an argument
'           : the same way BuildAndInstall receives its status file. The name carries a
'           : per-launch suffix, because the queue key cannot serve here: it is a hash of
'           : key and action, so it is identical in every session and forever, and a file
'           : left behind by an earlier run would be read as this answer.
'           : Returns an empty string if the folder cannot be prepared. The worker treats
'           : that as "do not write", and the caller reads it as no answer.
'---------------------------------------------------------------------------------------
'
Private Function JobResultFile(strAction As String) As String

    ' Well past the longest job timeout, so nothing in flight can match.
    Const StaleJobHours As Double = 1

    Dim oFS As Object
    Dim strFolder As String
    Dim colStale As Collection
    Dim varFile As Variant

    ' Dereference the FSO getter before opening a handler: it contains LogUnhandledErrors,
    ' so an error pending on entry would be reported from in here.
    Set oFS = FSO
    strFolder = oFS.GetParentFolderName(WorkerFilePath) & PathSep & "WorkerJobs"

    LogUnhandledErrors
    On Error Resume Next

    ' VerifyPath tells a folder from a file name by the trailing separator.
    VerifyPath strFolder & PathSep

    If oFS.FolderExists(strFolder) Then

        ' Nothing else collects these, because each name is unique to its launch. Collect
        ' first and delete afterwards: the Files collection is live.
        Set colStale = New Collection
        For Each varFile In oFS.GetFolder(strFolder).Files
            If varFile.DateLastModified < Now - (StaleJobHours / 24) Then
                colStale.Add varFile.Path
            End If
        Next varFile
        For Each varFile In colStale
            DeleteFile CStr(varFile)
        Next varFile

        JobResultFile = strFolder & PathSep & strAction & "-" & _
            UniqueHashSuffix(strAction) & ".result"
    End If

    CatchAny eelWarning, "Unable to prepare the worker result folder: " & strFolder, _
        ModuleName(Me) & ".JobResultFile"

End Function


'---------------------------------------------------------------------------------------
' Procedure : PollJobResults
' Author    : Adam Waller
' Date      : 8/23/2026
' Purpose   : Pick up a result delivered through the file channel and hand it to
'           : ReturnWorker, so that job logging and queue removal stay in one place.
'           : Anything other than the two values the worker writes is left alone and
'           : retried on the next interval: a read can fail transiently -- a scanner
'           : holding a file that was just published is the ordinary case -- and
'           : answering "not accessible" to that would be wrong where waiting one more
'           : tenth of a second is right. The job timeout already bounds the retrying.
'---------------------------------------------------------------------------------------
'
Private Sub PollJobResults()

    Dim oFS As Object
    Dim strContent As String
    Dim strValue As String

    If Len(m_strResultFile) = 0 Then Exit Sub

    Set oFS = FSO

    LogUnhandledErrors
    On Error Resume Next

    If oFS.FileExists(m_strResultFile) Then
        strContent = ReadFile(m_strResultFile)
        If Err.Number = 0 Then
            ' Split on an empty string returns an array with no elements.
            If Len(strContent) > 0 Then strValue = Trim(Split(strContent, vbCrLf)(0))
            Select Case strValue
                Case "1", "0"
                    Me.ReturnWorker m_strResultKey, (strValue = "1")
                    DeleteFile m_strResultFile
                    m_strResultFile = vbNullString
                    m_strResultKey = vbNullString
            End Select
        End If
        Err.Clear
    End If

    CatchAny eelWarning, "Unable to read the worker result file", _
        ModuleName(Me) & ".PollJobResults"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Initialize the worker queue.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set Me.Queue = New Dictionary
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReturnWorker
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Complete any post-call processes and remove from job queue.
'---------------------------------------------------------------------------------------
'
Public Function ReturnWorker(strKey As String, varParams As Variant)

    Dim cJob As clsJob

    ' Guard clause in case the code is stopped while workers are running.
    If Me.Queue Is Nothing Then Exit Function

    ' Make sure this job exists in the queue
    If Not Me.Queue.Exists(strKey) Then
        Log.Error eelError, "Returned worker not found in job queue: " & strKey, _
            ModuleName(Me) & ".ReturnWorker"
        Exit Function
    Else
        Set cJob = Me.Queue(strKey)
    End If

    ' Store the return value so the caller can read it after WaitForQueue
    m_varLastResult = varParams

    ' Log and remove from job queue
    With cJob
        Log.Add "Worker job " & .Action & " (" & .KeyHash & ") completed in " & _
            Round(Perf.MicroTimer - .Start, 2) & " seconds.", False
    End With
    Set cJob = Nothing
    Me.Queue.Remove strKey

End Function


'---------------------------------------------------------------------------------------
' *** BEGIN WORKER SCRIPT ***
'---------------------------------------------------------------------------------------
' WARNING: Everything below this marker is extracted verbatim by GetWorkerScriptContent
' and deployed as Worker.vbs (Windows Script Host), NOT VBA. It only compiles in the
' VBA IDE because of placeholder objects (e.g. Private WScript As Object).
'
' VBScript compatibility rules:
'   - Never skip optional arguments with ", ," (e.g. Replace(a, b, c, , 1) crashes in
'     VBScript with type mismatch 13; pass start/count explicitly). See issue #722.
'   - Do not use VBA-only constants or intrinsics that VBScript does not define.
'---------------------------------------------------------------------------------------
' Procedure : Main
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Main entry point for worker script
'---------------------------------------------------------------------------------------
'
Sub Main()

    Dim objFSO
    Dim objApp
    Dim strLibPath
    Dim strKey
    Dim strValue
    Dim varReturn
    Dim strUnattended
    Dim strStatusFile
    Dim strAction
    Dim strAttachError
    Dim strAccessPath
    Dim strLaunchPid
    Dim strPhaseStarted
    Dim strCallbackUrl
    Dim strOperationId
    Dim strResultFile


    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' We are expecting some very specific command line arguments in this worker script.
    If WScript.Arguments.Count < 3 Then
        MsgBox "Invalid arguments provided to worker script", vbExclamation
        WScript.Quit
    End If

    ' First argument is the path of the caller's current database, second the key,
    ' third the action.
    strValue = WScript.Arguments(0)
    strKey = WScript.Arguments(1)
    strAction = WScript.Arguments(2)

    ' Attach to the instance that launched this script. Every action but one is an
    ' in-process helper for a live database, and needs both the file and the reference.
    '
    ' BuildAndInstall is the exception, and must not even attempt this. GetObject binds
    ' a file moniker, which Access only honours for a database extension, so the bind
    ' fails whenever the caller's current database is the add-in itself -- and it can
    ' take twenty seconds to say so, which is time spent before this script can report
    ' that it started. Everything that action needs arrives in its own arguments, so it
    ' does not care what the caller had open, or whether it had anything open at all.
    '
    ' CheckDatabaseAccessible is the second, and for the same reason read the other way
    ' round: it builds its own DAO engine and opens the file itself, so the only thing it
    ' ever used the reference for was handing the answer back. That made the attach a
    ' dependency taken out purely in order to speak, and it is the one that cannot
    ' succeed for an add-in path. It reports through a result file instead, so its value
    ' no longer depends on reaching Access at all.
    Set objApp = Nothing
    If strAction <> "BuildAndInstall" And strAction <> "CheckDatabaseAccessible" Then

        If Not objFSO.FileExists(strValue) Then
            MsgBox "Database file must exist", vbExclamation
            WScript.Quit
        End If

        On Error Resume Next
        Set objApp = GetObject(strValue)
        strAttachError = "Error " & Err.Number & ": " & Err.Description
        Err.Clear
        If Not objApp Is Nothing Then
            Set objApp.VBE.ActiveVBProject = CurrentVBProject(objApp)
            Err.Clear
        End If
        On Error GoTo 0
        If objApp Is Nothing Then
            MsgBox "Could not attach to Microsoft Access for:" & vbCrLf & strValue & _
                vbCrLf & vbCrLf & strAttachError, vbExclamation
            WScript.Quit
        End If
    End If

    ' Determine specified action (More complex actions can be split out into subs)
    Select Case strAction

        Case "DeleteDatabaseObject"
            ' Type, Name
            objApp.DoCmd.DeleteObject WScript.Arguments(3), WScript.Arguments(4)

        Case "RenameDatabaseObject"
            ' New name, Type, Old Name
            objApp.DoCmd.Rename WScript.Arguments(3), WScript.Arguments(4), WScript.Arguments(5)

        Case "Uninstall"
            ' Remove the add-in database file and this worker script.
            ' (This must be done after closing Access)
            ' Args: Access application instance, add-in path
            UninstallAddIn objApp, WScript.Arguments(3)

        Case "BuildAndInstall"
            ' Rebuild the add-in from source and install the result.
            ' (This must be done after closing Access)
            ' Args: Access application instance, source path, unattended flag, status
            ' file, MSACCESS.EXE path, launching process id
            strUnattended = "0"
            strStatusFile = ""
            strAccessPath = ""
            strLaunchPid = ""
            strPhaseStarted = ""
            strCallbackUrl = ""
            strOperationId = ""
            If WScript.Arguments.Count >= 5 Then strUnattended = WScript.Arguments(4)
            If WScript.Arguments.Count >= 6 Then strStatusFile = WScript.Arguments(5)
            If WScript.Arguments.Count >= 7 Then strAccessPath = WScript.Arguments(6)
            If WScript.Arguments.Count >= 8 Then strLaunchPid = WScript.Arguments(7)
            If WScript.Arguments.Count >= 9 Then strPhaseStarted = WScript.Arguments(8)
            If WScript.Arguments.Count >= 10 Then strCallbackUrl = WScript.Arguments(9)
            If WScript.Arguments.Count >= 11 Then strOperationId = WScript.Arguments(10)
            BuildAndInstall WScript.Arguments(3), strUnattended, strStatusFile, _
                strAccessPath, strLaunchPid, strPhaseStarted, strCallbackUrl, _
                strOperationId

        Case "CheckDatabaseAccessible"
            ' Args: result file. The subject is the caller's current database, which
            ' arrives as argument 0 like everywhere else.
            '
            ' The file has to exist, and this action no longer passes through the branch
            ' that checks it. Guard it here instead, and on a miss write nothing: during
            ' a full build the database is renamed aside for a moment, and a path that
            ' resolves to nothing is not an answer. Writing a 0 would turn it into a
            ' definite "not accessible" and cost the caller a reopen it does not need.
            strResultFile = ""
            If WScript.Arguments.Count > 3 Then strResultFile = WScript.Arguments(3)
            If objFSO.FileExists(strValue) Then
                WriteJobResult strResultFile, CheckDatabaseAccessible(strValue)
            End If

        Case "SaveVbaProject"
            varReturn = SaveVbaProject(objApp)

        Case Else
            MsgBox "Unknown Action: " & strAction, vbExclamation

    End Select

    ' Make sure we still have a valid reference to the application before
    ' attempting any callback operations.
    If Not objApp Is Nothing Then

        ' Build path to library so we can ensure that we are calling
        ' the add-in function, not a function in the current database.
        strLibPath = GetAddInVBProject(objApp).FileName
        ' Remove file extension from library path.
        strLibPath = Left(strLibPath, Len(strLibPath) - Len(".accda"))

        ' Return key to caller to signal job completion
        objApp.Run strLibPath & ".WorkerCallback", strKey, varReturn
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SafeProjectFileName
' Author    : Adam Waller
' Date      : 6/9/2026
' Purpose   : Guarded read of a VBProject's FileName for the isolated worker script.
'           : The VBProjects collection can contain non-VBA entries (registered type
'           : libraries or wizards) that raise an error when .FileName is accessed,
'           : and the project set can be in flux while the database is reopened
'           : (e.g. in shared mode). Returns an empty string on failure. (#709)
'---------------------------------------------------------------------------------------
'
Function SafeProjectFileName(objProj)
    On Error Resume Next
    SafeProjectFileName = objProj.FileName
    On Error GoTo 0
End Function


'---------------------------------------------------------------------------------------
' Procedure : CurrentVBProject
' Author    : Adam Waller
' Date      : 3/3/2023
' Purpose   : Get a reference to the VB Project associated with the current database.
'           : (Look for a matching file name between the VB Projects and the database.)
'---------------------------------------------------------------------------------------
'
Function CurrentVBProject(objApp)

    Dim objProj
    Dim strUncPath

    ' Use currently active project by default
    Set CurrentVBProject = objApp.VBE.ActiveVBProject

    ' VBProject filenames are UNC paths
    strUncPath = GetUncPath(objApp.CurrentProject.FullName)

    If StrComp(SafeProjectFileName(objApp.VBE.ActiveVBProject), strUncPath, vbTextCompare) <> 0 Then
        ' Search for project with matching filename.
        For Each objProj In objApp.VBE.VBProjects
            If StrComp(SafeProjectFileName(objProj), strUncPath, vbTextCompare) = 0 Then
                Set CurrentVBProject = objProj
                Exit For
            End If
        Next
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAddInVBProject
' Author    : Adam Waller
' Date      : 3/3/2023
' Purpose   : Get a reference to the add-in project, even if it has the same name as
'           : the current project.
'---------------------------------------------------------------------------------------
'
Function GetAddInVBProject(objApp)

    Dim objProj
    Dim strUncPath
    Dim objFSO
    Dim strProjFile

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' VBProject filenames are UNC paths
    strUncPath = GetUncPath(objApp.CurrentProject.FullName)

    ' Search for project with matching filename.
    For Each objProj In objApp.VBE.VBProjects
        strProjFile = SafeProjectFileName(objProj)
        If Len(strProjFile) > 0 Then
            If StrComp(objFSO.GetBaseName(strProjFile), "Version Control", vbTextCompare) = 0 Then
                ' Check for different file path
                If StrComp(strProjFile, strUncPath, vbTextCompare) <> 0 Then
                    Set GetAddInVBProject = objProj
                    Exit For
                End If
            End If
        End If
    Next

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetUncPath
' Author    : Adam Waller
' Date      : 7/14/2020
' Purpose   : Returns the UNC path for a network location (if applicable)
'---------------------------------------------------------------------------------------
'
Function GetUncPath(strPath)

    Dim strDrive
    Dim strUNC
    Dim objFSO

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    strUNC = strPath
    strDrive = objFSO.GetDriveName(strPath)
    If strDrive <> vbNullString Then
        With objFSO.GetDrive(strDrive)
            If .DriveType = 3 Then  ' Remote
                ' VBScript requires start; VBA allows ", ," (issue #722).
                strUNC = Replace(strPath, strDrive, .ShareName, 1, 1, vbTextCompare)
            End If
        End With
    End If
    GetUncPath = strUNC

End Function


'---------------------------------------------------------------------------------------
' Procedure : UninstallAddIn
' Author    : Adam Waller
' Date      : 5/22/2023
' Purpose   : Remove add-in files that would normally be in use during the uninstall
'           : process. Then remove this worker script file itself.
'           : NOTE: No message is given if the uninstall succeeds.
'---------------------------------------------------------------------------------------
'
Function UninstallAddIn(objApp, strAddInFile)

    Dim objFSO
    Dim strLockFile
    Dim intCnt

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' Close the Microsoft Access application
    objApp.Quit
    Set objApp = Nothing

    ' Build lock file name from add-in name (.laccdb, not .laccda)
    strLockFile = Replace(strAddInFile, ".accda", ".laccdb", 1, 1, vbTextCompare)
    strLockFile = Replace(strLockFile, ".accde", ".laccdb", 1, 1, vbTextCompare)

    ' Resume next on any errors while attempting to delete the file.
    On Error Resume Next

    ' Wait up to 10 seconds to delete the main add-in file and database lock file.
    For intCnt = 1 To 10
        If objFSO.FileExists(strLockFile) Then objFSO.DeleteFile strLockFile
        If objFSO.FileExists(strAddInFile) Then objFSO.DeleteFile strAddInFile
        If Not objFSO.FileExists(strAddInFile) Then
            ' Clear any errors from failed delete attempts
            If Err Then Err.Clear
            Exit For
        End If
        WScript.Sleep 1000  ' Pause one second
    Next

    ' Ready or not, here we go with one final attempt. (This time we will show the error message.)
    If objFSO.FileExists(strAddInFile) Then objFSO.DeleteFile strAddInFile
    If Err Then
        MsgBox "Unable to delete add-in file: " & strAddInFile & vbCrLf & vbCrLf & _
            " The file might be in use by another instance of Microsoft Access." & vbCrLf & vbCrLf & _
            "Error " & Err.Number & ": " & Err.Description, vbExclamation
    Else
        ' Delete this worker script file.
        objFSO.DeleteFile WScript.ScriptFullName
        WScript.Quit
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckDatabaseAccessible
' Author    : Adam Waller
' Date      : 4/2/2026
' Purpose   : Test whether the database file can be opened by an external client.
'           : Uses a separate DAO engine instance (out-of-process) to detect the
'           : JET/ACE engine-level lock state (error 3734) that blocks connections
'           : after build/merge operations involving schema changes.
'---------------------------------------------------------------------------------------
'
Function CheckDatabaseAccessible(strPath)
    Dim objEngine
    Dim objDb
    On Error Resume Next
    Set objEngine = CreateObject("DAO.DBEngine.120")
    Set objDb = objEngine.OpenDatabase(strPath, False, True)
    If Err.Number = 0 Then
        objDb.Close
        Set objDb = Nothing
        CheckDatabaseAccessible = True
    Else
        CheckDatabaseAccessible = False
    End If
    Set objEngine = Nothing
End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteJobResult
' Author    : Adam Waller
' Date      : 8/23/2026
' Purpose   : Report a job's return value through a file, for an action that does not
'           : attach and so has no callback to make. Written as 1 or 0 rather than as a
'           : boolean so that nothing in the channel depends on locale, and with the same
'           : stream and charset as WriteRebuildStatusFile, which is the other file this
'           : script writes for the add-in to read back.
'           : Writes nothing when there is no path or no value. The absence of the file is
'           : how the caller tells "no answer" apart from "no", so writing anything here
'           : on a miss would destroy the distinction the file exists to make.
'---------------------------------------------------------------------------------------
'
Sub WriteJobResult(strFile, varValue)

    Dim stm

    If Len(strFile) = 0 Then Exit Sub
    If IsEmpty(varValue) Then Exit Sub

    On Error Resume Next
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    If varValue Then stm.WriteText "1" Else stm.WriteText "0"
    stm.SaveToFile strFile, 2
    stm.Close
    On Error GoTo 0

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveVbaProject
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Save the current database's VBA project by executing the VBE Save command
'           : out of process. The identical in-process call reports success and saves
'           : nothing; issued from here it saves. Main has already pointed ActiveVBProject
'           : at the current database's project. Returns the project's actual Saved state,
'           : not whether the command was issued.
'           :
'           : Do NOT reset the project here before saving. A reset ends whatever code is
'           : running, which during an export is the add-in waiting in WaitForQueue for this
'           : job -- so it terminates its own caller and empties the job queue. Tried and
'           : reverted; see modVbeUtility.SaveCurrentVBProject.
'---------------------------------------------------------------------------------------
'
Function SaveVbaProject(objApp)

    Dim objCtl
    Dim objProj

    On Error Resume Next

    Set objProj = CurrentVBProject(objApp)
    If objProj Is Nothing Then
        SaveVbaProject = False
        Exit Function
    End If

    If Not objProj.Saved Then
        Set objApp.VBE.ActiveVBProject = objProj
        ' ID 3 is the language-independent VBE Save command.
        Set objCtl = objApp.VBE.CommandBars.FindControl(, 3)
        If Not objCtl Is Nothing Then objCtl.Execute
    End If

    SaveVbaProject = objProj.Saved
    Set objCtl = Nothing
    Set objProj = Nothing

End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildAndInstall
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Rebuild the VCS add-in from source files and install the result.
'           : Waits for the launching Access instance to exit, opens a new one to build
'           : via the installed library, polls the build log for completion, then
'           : launches the install phase via /cmd INSTALL (or SILENT when unattended).
'           :
'           : Takes no Access reference: Main does not attach for this action. What
'           : used to be read back over COM -- the executable path -- and the means of
'           : closing the caller are arguments instead, so a rebuild works no matter
'           : what the caller had open.
'---------------------------------------------------------------------------------------
'
Function BuildAndInstall(strSourcePath, strUnattended, strStatusFile, _
    strAccessPath, strLaunchPid, strPhaseStarted, strCallbackUrl, strOperationId)

    Dim objFSO
    Dim objNewAccess
    Dim WshShell
    Dim strMSAccessPath
    Dim strInstalledLib
    Dim strRepoAddinPath
    Dim strLogsPath
    Dim strLockFile
    Dim intCnt
    Dim blnUnattended
    Dim strInstalledFile
    Dim dtBefore
    Dim strCmd
    Dim strCallbackInfo
    Dim strBuildResult
    Dim strInstallStatusFile

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set WshShell = CreateObject("WScript.Shell")

    blnUnattended = (strUnattended = "1")
    If Len(strPhaseStarted) = 0 Then strPhaseStarted = IsoDateTime(Now)

    ' Write this before anything that can fail. It is also the handshake the caller
    ' waits on, and the only proof available to it that this script survived its own
    ' startup. (See clsVersionControl.RebuildAddIn.)
    WriteRebuildStatusFile strStatusFile, "building", "", "", strPhaseStarted

    ' Supplied by the caller, which used to be read back over COM as objApp.SysCmd(9).
    ' Main no longer attaches for this action, so there is nothing to ask.
    strMSAccessPath = strAccessPath
    If Len(strMSAccessPath) = 0 Then
        FailRebuild Nothing, False, strStatusFile, blnUnattended, "build-failed", _
            "Could not determine the path of MSACCESS.EXE.", "", strPhaseStarted
        Exit Function
    End If

    ' Installed library path (Worker.vbs lives next to the installed .accda)
    strInstalledLib = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\Version Control"

    ' Repo .accda path: source folder is "Version Control.accda.src" so the
    ' .accda sits in the same parent directory.
    Dim strSrcClean
    strSrcClean = strSourcePath
    If Right(strSrcClean, 1) = "\" Then strSrcClean = Left(strSrcClean, Len(strSrcClean) - 1)
    strRepoAddinPath = objFSO.GetParentFolderName(strSrcClean) & "\Version Control.accda"

    ' Build log folder
    If Right(strSourcePath, 1) = "\" Then
        strLogsPath = strSourcePath & "logs\"
    Else
        strLogsPath = strSourcePath & "\logs\"
    End If

    ' RebuildAddIn closes its own instance, on a timer armed once it saw the status
    ' write above. Waiting on the process id is what proves the files it held are free,
    ' and unlike a COM quit it works whatever that instance had open.
    If Not WaitForProcessExit(strLaunchPid, 60) Then
        FailRebuild Nothing, False, strStatusFile, blnUnattended, "build-failed", _
            "The Access instance that launched the rebuild (process " & strLaunchPid & _
            ") did not exit, so the files it holds open could not be replaced.", _
            "", strPhaseStarted
        Exit Function
    End If

    ' Build lock file path for the repo .accda
    strLockFile = Replace(strRepoAddinPath, ".accda", ".laccdb", 1, 1, vbTextCompare)

    ' Wait for lock files to clear (up to 10 seconds)
    For intCnt = 1 To 10
        If Not objFSO.FileExists(strLockFile) Then Exit For
        WScript.Sleep 1000
    Next

    ' Snapshot existing build log files
    Dim dictExistingLogs
    Set dictExistingLogs = CreateObject("Scripting.Dictionary")
    SnapshotLogFiles objFSO, strLogsPath, "Build_", dictExistingLogs

    ' --- BUILD PHASE ---
    On Error Resume Next
    Set objNewAccess = CreateObject("Access.Application")
    If Err.Number <> 0 Then
        FailRebuild Nothing, False, strStatusFile, blnUnattended, "build-failed", _
            "Failed to create Access application instance." & vbCrLf & vbCrLf & _
            "Error " & Err.Number & ": " & Err.Description, "", strPhaseStarted
        Exit Function
    End If
    On Error GoTo 0

    ' Stage error trapping so expected On Error Resume Next blocks are not interrupted.
    ' eetBreakInClassModule = 1
    objNewAccess.SetOption "Error Trapping", 1

    objNewAccess.Visible = True
    objNewAccess.OpenCurrentDatabase strRepoAddinPath

    ' Close the installer form if AutoExec opened it
    On Error Resume Next
    objNewAccess.DoCmd.Close 2, "frmVCSInstall"  ' acForm = 2
    Err.Clear
    On Error GoTo 0

    ' MCP-launched builder: prefer a full-power core before the CPU-bound
    ' build. Missing on an older installed add-in; the MCP watcher also
    ' promotes new MSACCESS.EXE, so a failed Run is not fatal.
    If Len(strCallbackUrl) > 0 Then
        On Error Resume Next
        objNewAccess.Run strInstalledLib & ".PreferFullPowerCore"
        Err.Clear
        On Error GoTo 0
    End If

    ' Trigger build via the installed library. When an MCP callback identity was passed
    ' through the disconnected worker, use the same APIAsync entry point as ordinary
    ' database builds. This reuses Log.Add and Log.Progress rather than inventing a
    ' second stream. The worker still owns compile/install and their durable status.
    On Error Resume Next
    If Len(strCallbackUrl) > 0 And Len(strOperationId) > 0 Then
        strCallbackInfo = "{""callback_url"":""" & JsonEscape(strCallbackUrl) & _
            """,""operation_id"":""" & JsonEscape(strOperationId) & _
            """,""client"":""rebuild-worker""}"
        strBuildResult = objNewAccess.Run(strInstalledLib & ".APIAsync", _
            strCallbackInfo, "Build", strSourcePath)
        If InStr(1, strBuildResult, """async""", vbTextCompare) = 0 Then
            Err.Raise 1001, "BuildAndInstall", _
                "The asynchronous build entry point did not accept the rebuild."
        End If
    Else
        objNewAccess.Run strInstalledLib & ".HandleRibbonCommand", "btnBuild", strSourcePath
    End If
    If Err.Number <> 0 Then
        Dim strBuildErr
        strBuildErr = "Error " & Err.Number & ": " & Err.Description
        Err.Clear
        On Error GoTo 0
        FailRebuild objNewAccess, True, strStatusFile, blnUnattended, "build-failed", _
            "Failed to trigger build command." & vbCrLf & vbCrLf & strBuildErr, "", strPhaseStarted
        Exit Function
    End If
    On Error GoTo 0

    ' Poll for build log completion (2s interval, up to 120s)
    Dim strNewLog
    Dim strLogContent
    Dim intWait
    strNewLog = ""
    For intWait = 1 To 60
        WScript.Sleep 2000
        strNewLog = FindNewBuildLog(objFSO, strLogsPath, "Build_", dictExistingLogs)
        If Len(strNewLog) > 0 Then
            strLogContent = ReadTextFile(objFSO, strNewLog)
            If InStr(1, strLogContent, "TOTAL RUNTIME:", vbTextCompare) > 0 Then
                Exit For
            End If
            strNewLog = ""
        End If
    Next

    ' Check build result. The build instance is deliberately left running until the
    ' compile gate below has had its turn, since that gate needs the rebuilt project
    ' open, and a failing gate leaves the instance open for the developer.
    If Len(strNewLog) = 0 Then
        FailRebuild objNewAccess, True, strStatusFile, blnUnattended, "build-failed", _
            "Build timed out. No completed build log was found." & vbCrLf & vbCrLf & _
            "Check the logs folder for details:" & vbCrLf & strLogsPath, "", strPhaseStarted
        Exit Function
    End If

    ' Re-read the log to confirm success
    strLogContent = ReadTextFile(objFSO, strNewLog)
    Dim blnSuccess
    blnSuccess = (InStr(1, strLogContent, "Done. (", vbTextCompare) > 0) And _
                  (InStr(1, strLogContent, "CRITICAL:", vbTextCompare) = 0)

    If Not blnSuccess Then
        FailRebuild objNewAccess, True, strStatusFile, blnUnattended, "build-failed", _
            "Build failed. Check the build log for details:" & vbCrLf & vbCrLf & strNewLog, _
            strNewLog, strPhaseStarted
        Exit Function
    End If

    ' --- COMPILE GATE ---
    ' A build log only proves the source files imported. Compile before installing so
    ' code that does not build never reaches the installed add-in.
    WriteRebuildStatusFile strStatusFile, "compiling", "", strNewLog, strPhaseStarted
    If Not CompileBuiltAddIn(objNewAccess) Then
        On Error Resume Next
        objNewAccess.VBE.MainWindow.Visible = True
        On Error GoTo 0
        FailRebuild objNewAccess, False, strStatusFile, blnUnattended, "compile-failed", _
            "The rebuilt add-in does not compile, so it was not installed." & vbCrLf & vbCrLf & _
            strRepoAddinPath & vbCrLf & vbCrLf & _
            "Microsoft Access has been left open on the rebuilt file. Use " & _
            "Debug > Compile in the VBA editor to find the error, fix the source " & _
            "files, and run Rebuild Add-In again.", strNewLog, strPhaseStarted
        Exit Function
    End If

    ' Close the build Access instance
    QuitAccessInstance objNewAccess
    WScript.Sleep 2000

    ' --- INSTALL PHASE ---
    ' Wait for lock files from the build to clear
    For intCnt = 1 To 10
        If Not objFSO.FileExists(strLockFile) Then Exit For
        WScript.Sleep 1000
    Next
    If objFSO.FileExists(strLockFile) Then
        FailRebuild Nothing, False, strStatusFile, blnUnattended, "install-failed", _
            "The build Access instance did not release the rebuilt database before installation:" & _
            vbCrLf & strRepoAddinPath, strNewLog, strPhaseStarted
        Exit Function
    End If

    strInstalledFile = FindInstalledAddInFile(objFSO)
    dtBefore = #1/1/2000#
    If Len(strInstalledFile) > 0 Then
        If objFSO.FileExists(strInstalledFile) Then
            dtBefore = objFSO.GetFile(strInstalledFile).DateLastModified
        End If
    End If

    WriteRebuildStatusFile strStatusFile, "installing", "", strNewLog, strPhaseStarted

    If blnUnattended Then
        ' Keep the installer's own terminal status separate. It records "complete"
        ' immediately before DoCmd.Quit; publishing that to the rebuild status lets
        ' the MCP watcher return while the installer Access process is still open.
        strInstallStatusFile = ""
        If Len(strStatusFile) > 0 Then
            strInstallStatusFile = strStatusFile & ".install"
            On Error Resume Next
            If objFSO.FileExists(strInstallStatusFile) Then objFSO.DeleteFile strInstallStatusFile
            On Error GoTo 0
        End If
        If Len(strInstallStatusFile) > 0 Then
            strCmd = """" & strMSAccessPath & """ """ & strRepoAddinPath & _
                """ /cmd ""INSTALL SILENT " & strInstallStatusFile & """"
        Else
            strCmd = """" & strMSAccessPath & """ """ & strRepoAddinPath & """ /cmd ""INSTALL SILENT"""
        End If
        WshShell.Run strCmd, 7, True
        On Error Resume Next
        If Len(strInstallStatusFile) > 0 Then
            If objFSO.FileExists(strInstallStatusFile) Then objFSO.DeleteFile strInstallStatusFile
        End If
        On Error GoTo 0
    Else
        ' Launch Access with /cmd INSTALL (blocks until user clicks OK on success dialog)
        strCmd = """" & strMSAccessPath & """ """ & strRepoAddinPath & """ /cmd INSTALL"
        WshShell.Run strCmd, 1, True
    End If

    strInstalledFile = FindInstalledAddInFile(objFSO)
    If Len(strInstalledFile) = 0 Then
        FailRebuild Nothing, False, strStatusFile, blnUnattended, "install-failed", _
            "Install finished but the installed add-in file was not found.", strNewLog, strPhaseStarted
        Exit Function
    End If
    If objFSO.GetFile(strInstalledFile).DateLastModified <= dtBefore Then
        FailRebuild Nothing, False, strStatusFile, blnUnattended, "install-failed", _
            "Install finished but the installed add-in file was not updated:" & vbCrLf & strInstalledFile, _
            strNewLog, strPhaseStarted
        Exit Function
    End If

    WriteRebuildStatusFile strStatusFile, "complete", "", strNewLog, strPhaseStarted

End Function


'---------------------------------------------------------------------------------------
' Procedure : WaitForProcessExit
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Block until a process id is gone, and report whether it went. This is how
'           : the rebuild knows the launching instance released the files it held,
'           : without needing a COM reference to it.
'           : An unanswerable process query returns True rather than stalling the
'           : rebuild on a question Windows would not answer; the lock-file wait that
'           : follows is the second line of defense.
'---------------------------------------------------------------------------------------
'
Function WaitForProcessExit(strPid, intTimeoutSeconds)

    Dim objWmi
    Dim colProcs
    Dim intCnt

    ' Nothing to wait for
    If Len(strPid) = 0 Then
        WaitForProcessExit = True
        Exit Function
    End If

    Set objWmi = Nothing
    On Error Resume Next
    Set objWmi = GetObject("winmgmts:\\.\root\cimv2")
    Err.Clear
    On Error GoTo 0
    If objWmi Is Nothing Then
        WaitForProcessExit = True
        Exit Function
    End If

    For intCnt = 1 To intTimeoutSeconds
        Set colProcs = Nothing
        On Error Resume Next
        Set colProcs = objWmi.ExecQuery( _
            "SELECT ProcessId FROM Win32_Process WHERE ProcessId = " & strPid)
        Err.Clear
        On Error GoTo 0
        If colProcs Is Nothing Then
            WaitForProcessExit = True
            Exit Function
        End If
        If colProcs.Count = 0 Then
            WaitForProcessExit = True
            Exit Function
        End If
        WScript.Sleep 1000
    Next

End Function


'---------------------------------------------------------------------------------------
' Procedure : FailRebuild
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Record a terminal rebuild failure. Optionally quit the build instance.
'           : Shows a MsgBox only when attended. Compile-gate callers pass blnQuit=False
'           : so Access stays open on the rebuilt file.
'---------------------------------------------------------------------------------------
'
Sub FailRebuild(objApp, blnQuit, strStatusFile, blnUnattended, strStatus, strError, strBuildLog, strPhaseStarted)
    If blnQuit Then QuitAccessInstance objApp
    WriteRebuildStatusFile strStatusFile, strStatus, strError, strBuildLog, strPhaseStarted
    If Not blnUnattended Then
        MsgBox strError, vbCritical, "Rebuild Add-In"
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : WriteRebuildStatusFile
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Hand-rolled JSON writer for the worker script. No-ops when the path is
'           : empty. UTF-8 via ADODB.Stream so paths round-trip for the agent Read tool.
'---------------------------------------------------------------------------------------
'
Sub WriteRebuildStatusFile(strFile, strStatus, strError, strBuildLog, strPhaseStarted)
    Dim objFSO
    Dim strFolder
    Dim strJson
    Dim strUpdated
    Dim stm

    If Len(strFile) = 0 Then Exit Sub

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strFolder = objFSO.GetParentFolderName(strFile)
    If Len(strFolder) > 0 Then
        If Not objFSO.FolderExists(strFolder) Then
            On Error Resume Next
            objFSO.CreateFolder strFolder
            On Error GoTo 0
        End If
    End If

    strUpdated = IsoDateTime(Now)
    If Len(strPhaseStarted) = 0 Then strPhaseStarted = strUpdated

    strJson = "{" & vbCrLf & _
        "  ""status"": """ & JsonEscape(strStatus) & """," & vbCrLf & _
        "  ""error"": """ & JsonEscape(strError) & """," & vbCrLf & _
        "  ""buildLog"": """ & JsonEscape(strBuildLog) & """," & vbCrLf & _
        "  ""phaseStarted"": """ & JsonEscape(strPhaseStarted) & """," & vbCrLf & _
        "  ""updated"": """ & JsonEscape(strUpdated) & """" & vbCrLf & _
        "}"

    On Error Resume Next
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText strJson
    stm.SaveToFile strFile, 2
    stm.Close
    On Error GoTo 0
End Sub


'---------------------------------------------------------------------------------------
' Procedure : JsonEscape
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Escape a string for inclusion in a JSON value. Backslash first.
'---------------------------------------------------------------------------------------
'
Function JsonEscape(strValue)
    Dim strOut
    strOut = strValue
    If IsNull(strOut) Then strOut = ""
    strOut = Replace(strOut, "\", "\\")
    strOut = Replace(strOut, """", "\""")
    strOut = Replace(strOut, vbCrLf, "\n")
    strOut = Replace(strOut, vbCr, "\n")
    strOut = Replace(strOut, vbLf, "\n")
    JsonEscape = strOut
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsoDateTime
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : yyyy-mm-dd hh:nn:ss for status-file timestamps (VBScript has no Format$).
'---------------------------------------------------------------------------------------
'
Function IsoDateTime(dtValue)
    IsoDateTime = Year(dtValue) & "-" & Right("0" & Month(dtValue), 2) & "-" & _
        Right("0" & Day(dtValue), 2) & " " & Right("0" & Hour(dtValue), 2) & ":" & _
        Right("0" & Minute(dtValue), 2) & ":" & Right("0" & Second(dtValue), 2)
End Function


'---------------------------------------------------------------------------------------
' Procedure : FindInstalledAddInFile
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Return the installed add-in path next to Worker.vbs (.accda or .accde).
'---------------------------------------------------------------------------------------
'
Function FindInstalledAddInFile(objFSO)
    Dim strFolder
    Dim strAccda
    Dim strAccde
    strFolder = objFSO.GetParentFolderName(WScript.ScriptFullName)
    strAccda = strFolder & "\Version Control.accda"
    strAccde = strFolder & "\Version Control.accde"
    If objFSO.FileExists(strAccda) Then
        FindInstalledAddInFile = strAccda
    ElseIf objFSO.FileExists(strAccde) Then
        FindInstalledAddInFile = strAccde
    Else
        FindInstalledAddInFile = ""
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : QuitAccessInstance
' Author    : Adam Waller
' Date      : 7/31/2026
' Purpose   : Close an Access instance created by this script and release the reference.
'           : Quitting can fail when the instance is already gone, and every caller
'           : treats that the same way, so the error is swallowed here.
'---------------------------------------------------------------------------------------
'
Sub QuitAccessInstance(objApp)
    On Error Resume Next
    objApp.Quit
    On Error GoTo 0
    Set objApp = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CompileBuiltAddIn
' Author    : Adam Waller
' Date      : 7/31/2026
' Purpose   : Compile the freshly rebuilt add-in in the instance that built it, and
'           : return whether the project ended up in a compiled state. This is the gate
'           : that keeps code which does not compile from being installed: the build log
'           : only reports that source files imported, and the build itself never
'           : compiles (a project that does not compile still has to be buildable).
'           :
'           : The project is saved as well as compiled, so the installed copy ships
'           : compiled. Both commands are issued from out of process, which is where VBE
'           : and project-save operations have proven reliable for this add-in (see
'           : SaveVbaProject above, and modVbeUtility.SaveCurrentVBProject).
'---------------------------------------------------------------------------------------
'
Function CompileBuiltAddIn(objApp)

    Const VBE_CMD_RESET_ID = 228            ' Language-independent VBE Reset
    Const AC_CMD_COMPILE_AND_SAVE = 126     ' acCmdCompileAndSaveAllModules
    Const AC_CMD_COMPILE = 125              ' acCmdCompileAllModules

    Dim objProj
    Dim objCtl

    On Error Resume Next

    ' The build closed and recreated the database, so the project Main activated no
    ' longer exists. Compiling follows the active project, not the current database,
    ' so point the VBE at the rebuilt one before issuing the command.
    Set objProj = CurrentVBProject(objApp)
    If objProj Is Nothing Then Exit Function
    Set objApp.VBE.ActiveVBProject = objProj

    ' Clear run-state left behind by RunAfterBuild. Without this the compile can raise
    ' the modal "this action will reset your project" prompt, and nothing is present to
    ' dismiss it in a disconnected rebuild. Safe from here: no VBA is running in that
    ' instance, so the reset cannot end its own caller.
    Set objCtl = objApp.VBE.CommandBars.FindControl(, VBE_CMD_RESET_ID)
    If Not objCtl Is Nothing Then objCtl.Execute
    WScript.Sleep 1000

    ' Compile and save, so the installed copy ships compiled.
    objApp.DoCmd.RunCommand AC_CMD_COMPILE_AND_SAVE
    Err.Clear

    ' A save that is refused (2046) must not read as a compile failure, so fall back to
    ' a plain compile before deciding.
    If Not objApp.IsCompiled Then
        objApp.DoCmd.RunCommand AC_CMD_COMPILE
        Err.Clear
    End If

    CompileBuiltAddIn = objApp.IsCompiled

    Set objCtl = Nothing
    Set objProj = Nothing

End Function


'---------------------------------------------------------------------------------------
' Procedure : SnapshotLogFiles
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Collect existing log file names matching a prefix into a dictionary.
'---------------------------------------------------------------------------------------
'
Sub SnapshotLogFiles(objFSO, strFolder, strPrefix, dictOut)
    Dim objFile
    If Not objFSO.FolderExists(strFolder) Then Exit Sub
    For Each objFile In objFSO.GetFolder(strFolder).Files
        If Left(objFile.Name, Len(strPrefix)) = strPrefix Then
            dictOut.Add objFile.Name, True
        End If
    Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : FindNewBuildLog
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Find a build log file that wasn't in the original snapshot.
'---------------------------------------------------------------------------------------
'
Function FindNewBuildLog(objFSO, strFolder, strPrefix, dictExisting)
    Dim objFile
    FindNewBuildLog = ""
    If Not objFSO.FolderExists(strFolder) Then Exit Function
    For Each objFile In objFSO.GetFolder(strFolder).Files
        If Left(objFile.Name, Len(strPrefix)) = strPrefix Then
            If Not dictExisting.Exists(objFile.Name) Then
                FindNewBuildLog = objFile.Path
                Exit Function
            End If
        End If
    Next
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadTextFile
' Author    : Adam Waller
' Date      : 5/8/2026
' Purpose   : Read entire contents of a text file.
'---------------------------------------------------------------------------------------
'
Function ReadTextFile(objFSO, strPath)
    Dim objStream
    ReadTextFile = ""
    If Not objFSO.FileExists(strPath) Then Exit Function
    Set objStream = objFSO.OpenTextFile(strPath, 1, False)
    If Not objStream.AtEndOfStream Then ReadTextFile = objStream.ReadAll
    objStream.Close
    Set objStream = Nothing
End Function
