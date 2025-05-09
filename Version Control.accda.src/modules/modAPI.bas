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
' Use hard-coded enum index values to preserve user settings in existing projects
' if new values are added in the future.

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
    [_Last]
End Enum

' Sanitize levels used for sanitizing general and color elements in source files.
Public Enum eSanitizeLevel
    eslNone = 0     ' Don't sanitize anything.
    eslMinimal = 1  ' Sanitize minimal things like GUIDs.
    eslStandard = 2 ' Remove non-critical elements that cause VCS noise between builds.
    eslExtended = 3 ' Remove as much as possible. May have possible negative effects.
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
' Procedure : Preload
' Author    : Adam Waller
' Date      : 2/22/2025
' Purpose   : A simple generic function that does nothing in particular, but allows
'           : us to call it from external applications (i.e. Ribbon) to ensure that
'           : the add-in has been loaded into Microsoft Access.
'---------------------------------------------------------------------------------------
'
Public Function Preload()
    ' Hello, Add-In!
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
    Dim strLibName As String
    Dim strRunCmd As String

    ' Make sure the add-in is loaded.
    If Not AddinLoaded Then LoadVCSAddIn

    ' When running code from the add-in project itself, it gets a little
    ' tricky because both the add-in and the currentdb have the same VBProject name.
    ' This means we can't just call `Run "MSAccessVCS.*" because it will run in
    ' the local project instead of the add-in. We can resolve this by using the
    ' full path to the add-in library instead. (#593)
    Set projAddIn = GetAddInProject
    If RunningOnLocal Then
        strLibName = GetRunCmdAddInFullLibName
    Else
        strLibName = PROJECT_NAME
    End If

    ' See if we should run the command directly, or with an API timer callback.
    ' (The API timer is helpful when you need to clear the call stack on the
    '  current database before running the add-in code.)
    If blnUseTimer And Not RunningOnLocal Then
        If Operation.Status = eosRunning Then Operation.Stage
        SetTimer strProcedure, CStr(varArg1), CStr(varArg2)
    Else
        ' Build the command to execute using Application.Run
        strRunCmd = strLibName & "." & strProcedure
        ' Call based on arguments
        If Not IsMissing(varArg2) Then
            Application.Run strRunCmd, varArg1, varArg2
        ElseIf Not IsMissing(varArg1) Then
            Application.Run strRunCmd, varArg1
        Else
            Application.Run strRunCmd
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetRunCmdAddInFullLibName
' Author    : Josef Poetzl
' Date      : 2/20/2025
' Purpose   : Return the full path to the add-in library without the file extension.
'---------------------------------------------------------------------------------------
'
Private Function GetRunCmdAddInFullLibName() As String

   Const cstrAddInFileExtension As String = ".accda"
   Dim strAddInFileName As String

   strAddInFileName = GetAddInFileName
   GetRunCmdAddInFullLibName = Left(strAddInFileName, Len(strAddInFileName) - Len(cstrAddInFileExtension))

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExampleLoadAddInAndRunExport
' Author    : Adam Waller
' Date      : 2/21/2025
' Purpose   : This function can be copied to a local database and triggered with a
'           : command line argument or other automation technique to load the VCS
'           : add-in file and initiate an export.
'           : NOTE: This expects the add-in to be installed in the default location
'           : and using the default file name.
'---------------------------------------------------------------------------------------
'
Public Function ExampleLoadAddInAndRunExport()
    Application.Run Environ$("AppData") & "\MSAccessVCS\Version Control" & _
        ".HandleRibbonCommand", "btnExport"
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExampleBuildFromSource
' Author    : Adam Waller
' Date      : 2/21/2025
' Purpose   : This function can be copied to a local database and triggered with a
'           : command line argument or other automation technique to load the VCS
'           : add-in file and build this project from source.
'           : NOTE: This expects the add-in to be installed in the default location
'           : and using the default file name.
'---------------------------------------------------------------------------------------
'
Public Function ExampleBuildFromSource(Optional strSourcePath As String)
    Application.Run Environ$("AppData") & "\MSAccessVCS\Version Control" & _
        ".HandleRibbonCommand", "btnBuild", strSourcePath
End Function
