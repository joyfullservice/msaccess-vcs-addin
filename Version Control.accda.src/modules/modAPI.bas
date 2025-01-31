Attribute VB_Name = "modAPI"
'---------------------------------------------------------------------------------------
' Module    : modAPI
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : This module exposes a set of VCS tools to other projects.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

' Note, some enums are listed here when they are directly exposed
' through the Options and VCS classes. (Allowing them to be used externally)

' Control the interaction mode
Public Enum eInteractionMode
    eimNormal = 0
    eimSilent = 1
End Enum

' Formats used when exporting table data.
Public Enum eTableDataExportFormat
    etdNoData = 0
    etdTabDelimited = 1
    etdXML = 2
    [_Last] = 2
End Enum

' Sanitize levels used for sanitizing general and color elements in source files.
Public Enum eSanitizeLevel
    eslNone = 0     ' Don't sanitize anything.
    eslMinimal = 1  ' Sanitize minimal things like GUIDs.
    eslStandard     ' Remove non-critical elements that cause VCS noise between builds.
    eslExtended     ' Remove as much as possible. May have possible negative effects.
    [_Last]         ' Placeholder for the end of the list.
End Enum


'---------------------------------------------------------------------------------------
' Procedure : HandleRibbonCommand
' Author    : Adam Waller
' Date      : 2/28/2022
' Purpose   : Handle an incoming command from the TwinBasic ribbon COM add-in.
'           : (Allows us to keep all the logic between the XML ribbon file and the
'           :  Access add-in.)
'---------------------------------------------------------------------------------------
'
Public Function HandleRibbonCommand(strCommand As String, Optional strArgument As String) As Boolean
    ' The function is called by Application.Run which can be re-entrant but we really
    ' don't want it to be since that'd cause errors. To avoid this, we will ignore any
    ' commands while the current command is running.
    Static IsRunning As Boolean

    On Error GoTo ErrHandler

    If IsRunning Then
        ' Ignore the re-entry; do NOT go to clean-up.
        Exit Function
    End If

    IsRunning = True

    ' Make sure we are not attempting to run this from the current database when making
    ' changes to the add-in itself. (It will re-run the command through the add-in.)
    If RunningOnLocal() Then
        RunInAddIn "HandleRibbonCommand", True, strCommand, strArgument
        GoTo CleanUp
    End If

    ' If a function is not found, this will throw an error. It is up to the ribbon
    ' designer to ensure that the control IDs match public procedures in the VCS
    ' (clsVersionControl) class module.
    ' For example, to run VCS.Export, the ribbon button ID should be named "btnExport"

    ' Trim off control ID prefix when calling command
    If Len(strArgument) Then
        CallByName VCS, Mid(strCommand, 4), VbMethod, strArgument
    Else
        CallByName VCS, Mid(strCommand, 4), VbMethod
    End If

CleanUp:
    IsRunning = False
    Exit Function

ErrHandler:
    ' An error occurred so we need to make it available for further attempts
    ' but do not handle the error.
    IsRunning = False

    ' Re-throw
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function


'---------------------------------------------------------------------------------------
' Procedure : VCS
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Wrapper for the VCS class, providing easy API access to VCS functions.
'           : *NOTE* that this class is not persisted. This allows us to wrap up and
'           : remove any object references after the call completes.
'---------------------------------------------------------------------------------------
'
Public Function VCS() As clsVersionControl
    Set VCS = New clsVersionControl
End Function


'---------------------------------------------------------------------------------------
' Procedure : WorkerCallback
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : A public callback endpoint for worker script operations to check back in
'           : after completion.
'---------------------------------------------------------------------------------------
'
Public Function WorkerCallback(strKey As String, Optional varParams As Variant)
    If RunningOnLocal Then
        RunInAddIn "WorkerCallback", False, strKey, varParams
    Else
        Worker.ReturnWorker strKey, varParams
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetInteractionMode
' Author    : Adam Waller
' Date      : 6/24/2023
' Purpose   : Control the types of UI interaction. (For example, you might set the
'           : interaction mode to silent during an automated build.)
'---------------------------------------------------------------------------------------
'
Public Sub SetInteractionMode(intMode As eInteractionMode)
    modVCSUtility.InteractionMode = intMode
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunningOnLocal
' Author    : Adam Waller
' Date      : 5/18/2022
' Purpose   : Returns true if the code is running in the current database instead of
'           : the add-in database.
'---------------------------------------------------------------------------------------
'
Private Function RunningOnLocal() As Boolean
    RunningOnLocal = (StrComp(CurrentProject.FullName, CodeProject.FullName, vbTextCompare) = 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : RunInAddIn
' Author    : Adam Waller
' Date      : 3/3/2023
' Purpose   : Run a proceedure with optional parameters in the VCS add-in database.
'---------------------------------------------------------------------------------------
'
Public Function RunInAddIn(strProcedure As String, blnUseTimer As Boolean, Optional varArg1 As Variant, Optional varArg2 As Variant)

    Dim projAddIn As VBProject
    Dim strRunCmd As String

    ' Make sure the add-in is loaded.
    If Not AddinLoaded Then LoadVCSAddIn

    ' When running code from the add-in project itself, it gets a little
    ' tricky because both the add-in and the currentdb have the same VBProject name.
    ' This means we can't just call `Run "MSAccessVCS.*" because it will run in
    ' the local project instead of the add-in. To pull this off, we will temporarily
    ' change the project name of the add-in so we can call it as distinct from the
    ' current project.
    Set projAddIn = GetAddInProject
    If RunningOnLocal Then
        ' When this is run from the CurrentDB, we should rename the add-in project,
        ' then call it again using the renamed project to ensure we are running it
        ' from the add-in.
        projAddIn.Name = "MSAccessVCS-Lib"
    Else
        ' Running from the add-in project
        ' Reset project name if needed
        If projAddIn.Name = "MSAccessVCS-Lib" Then projAddIn.Name = PROJECT_NAME
    End If

    ' See if we should run the command directly, or with an API timer callback.
    ' (The API timer is helpful when you need to clear the call stack on the
    '  current database before running the add-in code.)
    If blnUseTimer And Not RunningOnLocal Then
        SetTimer strProcedure, CStr(varArg1), CStr(varArg2)
    Else
        ' Build the command to execute using Application.Run
        strRunCmd = projAddIn.Name & "." & strProcedure
        ' Call based on arguments
        If Not IsMissing(varArg2) Then
            Run strRunCmd, varArg1, varArg2
        ElseIf Not IsMissing(varArg1) Then
            Run strRunCmd, varArg1
        Else
            Run strRunCmd
        End If
    End If

    ' Restore project name after run (if needed)
    If projAddIn.Name = "MSAccessVCS-Lib" Then projAddIn.Name = PROJECT_NAME

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExampleLoadAddInAndRunExport
' Author    : Adam Waller
' Date      : 11/13/2020
' Purpose   : This function can be copied to a local database and triggered with a
'           : command line argument or other automation technique to load the VCS
'           : add-in file and initiate an export.
'           : NOTE: This expects the add-in to be installed in the default location
'           : and using the default file name.
'---------------------------------------------------------------------------------------
'
Public Function ExampleLoadAddInAndRunExport()

    Dim strAddInPath As String
    Dim proj As Object      ' VBProject
    Dim objAddIn As Object  ' VBProject

    ' Build default add-in path
    strAddInPath = Environ$("AppData") & "\MSAccessVCS\Version Control.accda"

    ' See if add-in project is already loaded.
    For Each proj In VBE.VBProjects
        If StrComp(proj.FileName, strAddInPath, vbTextCompare) = 0 Then
            Set objAddIn = proj
        End If
    Next proj

    ' If not loaded, then attempt to load the add-in.
    If objAddIn Is Nothing Then

        ' The following lines will load the add-in at the application level,
        ' but will not actually call the function. Ignore the error of function not found.
        ' https://stackoverflow.com/questions/62270088/how-can-i-launch-an-access-add-in-not-com-add-in-from-vba-code
        On Error Resume Next
        Application.Run strAddInPath & "!DummyFunction"
        On Error GoTo 0

        ' See if it is loaded now...
        For Each proj In VBE.VBProjects
            If StrComp(proj.FileName, strAddInPath, vbTextCompare) = 0 Then
                Set objAddIn = proj
            End If
        Next proj
    End If

    If objAddIn Is Nothing Then
        MsgBox "Unable to load Version Control add-in. Please ensure that it has been installed" & vbCrLf & _
            "and is functioning correctly. (It should be available in the Add-ins menu.)", vbExclamation
    Else
        ' Launch add-in export for current database.
        Application.Run "MSAccessVCS.ExportSource", True
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExampleBuildFromSource
' Author    : Adam Waller
' Date      : 9/6/2023
' Purpose   : This function can be copied to a local database and triggered with a
'           : command line argument or other automation technique to load the VCS
'           : add-in file and build this project from source.
'           : NOTE: This expects the add-in to be installed in the default location
'           : and using the default file name.
'---------------------------------------------------------------------------------------
'
Public Function ExampleBuildFromSource()

    Dim strAddInPath As String
    Dim proj As Object      ' VBProject
    Dim objAddIn As Object  ' VBProject

    ' Build default add-in path
    strAddInPath = Environ$("AppData") & "\MSAccessVCS\Version Control.accda"

    ' See if add-in project is already loaded.
    For Each proj In VBE.VBProjects
        If StrComp(proj.FileName, strAddInPath, vbTextCompare) = 0 Then
            Set objAddIn = proj
        End If
    Next proj

    ' If not loaded, then attempt to load the add-in.
    If objAddIn Is Nothing Then

        ' The following lines will load the add-in at the application level,
        ' but will not actually call the function. Ignore the error of function not found.
        ' https://stackoverflow.com/questions/62270088/how-can-i-launch-an-access-add-in-not-com-add-in-from-vba-code
        On Error Resume Next
        Application.Run strAddInPath & "!DummyFunction"
        On Error GoTo 0

        ' See if it is loaded now...
        For Each proj In VBE.VBProjects
            If StrComp(proj.FileName, strAddInPath, vbTextCompare) = 0 Then
                Set objAddIn = proj
            End If
        Next proj
    End If

    If objAddIn Is Nothing Then
        MsgBox "Unable to load Version Control add-in. Please ensure that it has been installed" & vbCrLf & _
            "and is functioning correctly. (It should be available in the Add-ins menu.)", vbExclamation
    Else
        ' Set the application interaction level to silent to skip confirmation dialogs.
        Application.Run "MSAccessVCS.SetInteractionMode", 1
        ' Launch the build process (as if we clicked the button on the ribbon)
        ' Optionally specify a specific folder of source files to build from.
        Application.Run "MSAccessVCS.HandleRibbonCommand", "btnBuild" ', "c:\path\to\source\folder"
    End If

End Function
