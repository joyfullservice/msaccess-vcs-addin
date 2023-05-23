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
' Procedure : Run_UninstallAddin
' Author    : Adam Waller
' Date      : 5/22/2023
' Purpose   : Remove the actual add-in file and lock file during uninstall.
'---------------------------------------------------------------------------------------
'
Public Function Run_UninstallAddin()
    CallWorker PROJECT_NAME, 10, "Uninstall", """" & GetAddInFileName & """"
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
            ' Normally we don't want to stay on a constant loop
            ' of doevents, but in this case it is necessary to allow
            ' the worker callbacks to process.
            DoEvents
        Loop
        DoEvents

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
    Dim objProj
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
        'Set objApp.VBE.ActiveVBProject = GetAddInVBProject(objApp)
        Set objApp.VBE.ActiveVBProject = CurrentVBProject(objApp)
    End If

    ' The second argument should be the key
    strKey = WScript.Arguments(1)

    ' Determine specified action (More complex actions can be split out into subs)
    Select Case WScript.Arguments(2)

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

        Case Else
            MsgBox "Unknown Action: " & WScript.Arguments(1), vbExclamation

    End Select

    ' Make sure we still have a valid reference to the application before
    ' attempting any callback operations.
    If Not objApp Is Nothing Then
    
        ' Check the project name, just in case we are editing the VCS project
        If objApp.CurrentProject.Name = "Version Control.accda" Then
            ' Rename project so we can run the callback directly on the add-in
            ' even if the local project does not currently compile. (Such as
            ' after removing a class module during a merge operation.)
            Set objProj = GetAddInVBProject(objApp)
            objProj.Name = "MSAccessVCS-Lib"
            objApp.Run "MSAccessVCS-Lib.WorkerCallback", strKey, varReturn
            objProj.Name = "MSAccessVCS"
        Else
            ' Return key to caller to signal job completion
            objApp.Run "MSAccessVCS.WorkerCallback", strKey, varReturn
        End If
    End If

End Sub


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

    If StrComp(objApp.VBE.ActiveVBProject.FileName, strUncPath, vbTextCompare) <> 0 Then
        ' Search for project with matching filename.
        For Each objProj In objApp.VBE.VBProjects
            If StrComp(objProj.FileName, strUncPath, vbTextCompare) = 0 Then
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

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ' VBProject filenames are UNC paths
    strUncPath = GetUncPath(objApp.CurrentProject.FullName)

    ' Search for project with matching filename.
    For Each objProj In objApp.VBE.VBProjects
        If StrComp(objFSO.GetFileName(objProj.FileName), "Version Control.accda", vbTextCompare) = 0 Then
            ' Check for different file path
            If StrComp(objProj.FileName, strUncPath, vbTextCompare) <> 0 Then
                Set GetAddInVBProject = objProj
                Exit For
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
                strUNC = Replace(strPath, strDrive, .ShareName, , 1, vbTextCompare)
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
    
    ' Build lock file name from add-in name
    strLockFile = Replace(strAddInFile, ".accda", ".laccda")
    
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
