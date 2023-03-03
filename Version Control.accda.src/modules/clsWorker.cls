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
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


' Worker Job queue of jobs in progress
Public Queue As Dictionary

' Generic placeholder objects so we can compile the VBScript in the VBA IDE
Private WScript As Object


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
Public Function DeleteDatabaseObject(intType As AcObjectType, strName As String)

    ' Don't use the following method from the add-in! Instead, we need to call
    ' it from an external process so it deletes from the current database first.
    'DoCmd.DeleteObject acModule, "basModule"
    
    CallWorker strName, 5, "DeleteDatabaseObject", intType, strName
    
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
    Dim colRemove As Collection
    
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
        Pause sngPollSeconds
        DoEvents
        
        ' We can exit early if there is nothing in the queue
        If Me.Queue.Count = 0 Then Exit Do
            
        ' Reset flag for this iteration
        blnFound = False
        
        ' Loop through queue of jobs
        For Each varKey In Me.Queue.Keys
            Set cJob = Me.Queue.Item(varKey)
            
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
        
        ' Log and remove any individual jobs that have timed out
        
        
        ' If nothing found of a specified action, we can exit the loop
        If (strAction <> "All") And (Not blnFound) Then Exit Do
        
        ' Check the overall timeout before we start another iteration
        If dblStart + dblTimeout > Perf.MicroTimer Then
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
' Procedure : CallWorker
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Call the worker script to run a specified operation.
'---------------------------------------------------------------------------------------
'
Private Function CallWorker(strKey As String, dblTimeoutSeconds As Double, _
    strAction As String, ParamArray varParams()) As Long

    Dim intParam As Integer
    Dim strCommand As String
    Dim strHash As String
    Dim lngHandle As Long
    Dim cJob As clsJob
    
    ' Make sure this key/action is not already running
    strHash = GetStringHash(strKey & strAction)
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
        
        ' Use the Shell command to launch the worker as an external asynchrous
        ' process, and return the process ID to the caller.
        lngHandle = Shell(.GetStr, vbHide)
        
    End With
    
    ' Add to job queue
    If lngHandle <> 0 Then
        Set cJob = New clsJob
        With cJob
            .KeyHash = strHash
            .Action = strAction
            .Start = Perf.MicroTimer
            .Timeout = .Start + dblTimeoutSeconds
        End With
    End If
    
    ' Return process handle
    CallWorker = lngHandle
    
End Function


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

    ' Special processing for some actions
    Select Case cJob.Action
        Case Else
            ' No special processing
    End Select
    
    ' Log and remove from job queue
    With cJob
        Log.Add "Worker job " & .Action & " (" & .KeyHash & ") completed in " & _
            Round(Perf.MicroTimer - .Start, 2) & " seconds."
    End With
    Set cJob = Nothing
    Me.Queue.Remove strKey
    
End Function


'---------------------------------------------------------------------------------------
' *** BEGIN WORKER SCRIPT ***
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
    Dim strKey
    Dim strValue
    Dim varReturn
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' We are expecting some very specific command line arguments in this worker script.
    If WScript.Arguments.Count < 2 Then
        MsgBox "Invalid arguments provided to worker script", vbExclamation
        WScript.Quit
    End If

    ' The first argument should be the path of the current database
    strValue = WScript.Arguments(0)
    If Not objFSO.FileExists(strValue) Then
        MsgBox "Database file must exist", vbExclamation
        WScript.Quit
    Else
        ' Get a reference to the database application
        Set objApp = GetObject(strValue)
    End If
    
    ' The second argument should be the key
    strKey = WScript.Arguments(1)
    
    ' Determine specified action (More complex actions can be split out into subs)
    Select Case WScript.Arguments(2)
        
        Case "DeleteDatabaseObject"
            ' Type, Name
            objApp.DoCmd.DeleteObject WScript.Arguments(2), WScript.Arguments(3)
    
        Case "RenameDatabaseObject"
            ' New name, Type, Old Name
            objApp.DoCmd.DeleteObject WScript.Arguments(2), WScript.Arguments(3), WScript.Arguments(3)
        
        Case Else
            MsgBox "Unknown Action: " & WScript.Arguments(1), vbExclamation
            
    End Select

    ' Return key to caller to signal job completion
    objApp.Run "MSAccessVCS.WorkerCallback", strKey, varReturn
    
End Sub

