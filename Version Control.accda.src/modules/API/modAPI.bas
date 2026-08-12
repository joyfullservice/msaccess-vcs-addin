Attribute VB_Name = "modAPI"
'---------------------------------------------------------------------------------------
' Module    : modAPI
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : This module exposes a set of VCS tools to other projects.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("API")

' MCP Debug log file path (set when first used)
Private m_strMCPDebugLogPath As String

' Returned by API (and embedded in the APIAsync JSON) when a call arrives while another
' is still in progress. Published as a prefix rather than an error number because the
' refusal cannot be raised -- see RefuseReentrantCall. External callers (the MCP server
' in particular) match on this to report a refusal instead of treating it as data.
Public Const API_REFUSED_PREFIX As String = "VCS_API_REFUSED: "


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

' Controls whether connection strings with credentials are stored in a .env file.
' Export: Auto = only when UID/PWD detected; Always = all connections; Never = keep in source.
' Build: Auto = prompt to save new credentials; Always = save silently; Never = never write .env.
Public Enum eUseEnvConnections
    uecAuto = 0     ' Export when credentials detected; build prompts before saving.
    uecAlways = 1   ' Always store connection strings in .env file.
    uecNever = 2    ' Leave connection strings in source; never write .env on build.
    [_Last]
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

    LogUnhandledErrors
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

    ' Ribbon commands are interactive by definition. Operation.Source is sticky,
    ' so clear any value left behind by an earlier API/MCP call in this instance.
    Operation.Source = eosUserInterface

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
' Procedure : HandleTestAssertion
' Author    : Adam Waller
' Date      : 5/7/2026
' Purpose   : Entry point for TestAssert calls from user projects. Must live in a
'           : standard module so Application.Run can reach it.
'---------------------------------------------------------------------------------------
'
Public Function HandleTestAssertion(ByVal blnCondition As Boolean, _
                                     Optional ByVal varContext As Variant) As Boolean
    If TestRunner.State <> etrsRunning Then
        HandleTestAssertion = False
        Exit Function
    End If
    HandleTestAssertion = True
    TestRunner.RecordAssertion blnCondition, varContext
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
    Operation.InteractionMode = intMode
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

    Dim strLibName As String
    Dim strRunCmd As String

    ' Make sure the add-in is loaded.
    If Not AddinLoaded Then LoadVCSAddIn

    ' When running code from the add-in project itself, it gets a little
    ' tricky because both the add-in and the currentdb have the same VBProject name.
    ' This means we can't just call `Run "MSAccessVCS.*" because it will run in
    ' the local project instead of the add-in. We can resolve this by using the
    ' full path to the add-in library instead. (#593)
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
    GetRunCmdAddInFullLibName = FSO.BuildPath( _
        GetInstallSettings.strInstallFolder, ADDIN_BASENAME)
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
    Application.Run Environ$("AppData") & "\" & PROJECT_NAME & "\" & ADDIN_BASENAME & _
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
    Application.Run Environ$("AppData") & "\" & PROJECT_NAME & "\" & ADDIN_BASENAME & _
        ".HandleRibbonCommand", "btnBuild", strSourcePath
End Function


'---------------------------------------------------------------------------------------
' Procedure : API
' Author    : Adam Waller
' Date      : 1/22/2026
' Purpose   : Public API function for external automation (e.g., MCP tools) to call
'           : VCS methods via Application.Run and receive return values.
'           : Routes method calls to clsVersionControl class methods.
'           : Usage: Application.Run("Version Control.API", "GetVCSVersion")
'---------------------------------------------------------------------------------------
'
Public Function API(strMethod As String, _
    Optional varArg1 As Variant, _
    Optional varArg2 As Variant, _
    Optional varArg3 As Variant) As Variant

    ' The function is called by Application.Run, which can be re-entrant. We really don't
    ' want it to be, since a nested call would run against Operation state the outer call
    ' owns. Refuse it, and say so in the return value rather than coming back Empty --
    ' see RefuseReentrantCall for why silence was worse and why this is not an Err.Raise.
    Static IsRunning As Boolean
    Dim varResult As Variant
    Dim strLibName As String
    Dim strRunCmd As String

    SuppressErrorBreaks
    LogUnhandledErrors
    On Error GoTo ErrHandler

    If IsRunning Then
        API = RefuseReentrantCall(strMethod, "API")
        RestoreErrorBreaks
        Exit Function
    End If

    IsRunning = True

    ' Set operation source before any operations
    Operation.Source = eosExternalAPI

    ' Make sure we are not attempting to run this from the current database when making
    ' changes to the add-in itself. (It will re-run the command through the add-in.)
    If RunningOnLocal() Then
        ' When running from within the add-in database, we need to use the full path
        ' to ensure we call the add-in version, not a local version.
        strLibName = GetRunCmdAddInFullLibName
        strRunCmd = strLibName & ".API"

        ' Use Application.Run to call the add-in version, which will return the value
        If Not IsMissing(varArg3) Then
            API = Application.Run(strRunCmd, strMethod, varArg1, varArg2, varArg3)
        ElseIf Not IsMissing(varArg2) Then
            API = Application.Run(strRunCmd, strMethod, varArg1, varArg2)
        ElseIf Not IsMissing(varArg1) Then
            API = Application.Run(strRunCmd, strMethod, varArg1)
        Else
            API = Application.Run(strRunCmd, strMethod)
        End If
        GoTo CleanUp
    End If

    ' Use the VCS() function to get an instance (same pattern as HandleRibbonCommand)
    ' This ensures we don't interfere with any staging of settings during a running operation
    ' Use CallByName to invoke the method dynamically
    ' Handle different numbers of arguments
    ' Every parameter of a method reachable from here must be declared ByVal. These
    ' arguments arrive as Variants, and CallByName cannot bind a Variant to a ByRef typed
    ' parameter: it raises a type mismatch before the method runs a single line, which
    ' reads to the caller as the method having done nothing.
    If Not IsMissing(varArg3) Then
        ' Three arguments
        varResult = CallByName(VCS, strMethod, VbMethod, varArg1, varArg2, varArg3)
    ElseIf Not IsMissing(varArg2) Then
        ' Two arguments
        varResult = CallByName(VCS, strMethod, VbMethod, varArg1, varArg2)
    ElseIf Not IsMissing(varArg1) Then
        ' One argument
        varResult = CallByName(VCS, strMethod, VbMethod, varArg1)
    Else
        ' No arguments
        varResult = CallByName(VCS, strMethod, VbMethod)
    End If

    ' Return the result (will be Empty for Subs)
    API = varResult

CleanUp:
    IsRunning = False
    RestoreErrorBreaks
    Exit Function

ErrHandler:
    ' An error occurred so we need to make it available for further attempts
    ' but do not handle the error.
    IsRunning = False
    RestoreErrorBreaks

    ' Re-throw
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function


'---------------------------------------------------------------------------------------
' Procedure : APIAsync
' Author    : Adam Waller
' Date      : 1/23/2026
' Purpose   : Public API function for async operations with MCP callbacks.
'           : Routes async operations through timer mechanism with callback support.
'           : Usage: Application.Run("Version Control.APIAsync", strCallbackInfo, "Export")
'---------------------------------------------------------------------------------------
'
Public Function APIAsync(strCallbackInfo As String, strMethod As String, _
    Optional varArg1 As Variant, Optional varArg2 As Variant) As String

    ' The function is called by Application.Run, which can be re-entrant. We really don't
    ' want it to be, since a nested call would run against Operation state the outer call
    ' owns. Refuse it, and say so in the return value rather than coming back Empty --
    ' see RefuseReentrantCall for why silence was worse and why this is not an Err.Raise.
    Static IsRunning As Boolean
    Dim varResult As Variant
    Dim strLibName As String
    Dim strRunCmd As String
    Dim strJsonResult As String
    Dim dResult As Dictionary
    Dim lngTimeoutMs As Long

    SuppressErrorBreaks
    LogUnhandledErrors
    On Error GoTo ErrHandler

    If IsRunning Then
        ' Callers parse this return value as JSON, so the refusal has to arrive as JSON.
        Set dResult = New Dictionary
        dResult.Add "success", False
        dResult.Add "error", RefuseReentrantCall(strMethod, "APIAsync")
        APIAsync = modJsonConverter.ConvertToJson(dResult)
        RestoreErrorBreaks
        Exit Function
    End If

    IsRunning = True

    ' Set operation source - MCP if callback provided, otherwise external API
    If Len(strCallbackInfo) > 0 Then
        Operation.Source = eosMCPTool
    Else
        Operation.Source = eosExternalAPI
    End If

    ' Make sure we are not attempting to run this from the current database when making
    ' changes to the add-in itself. (It will re-run the command through the add-in.)
    If RunningOnLocal() Then
        ' When running from within the add-in database, we need to use the full path
        ' to ensure we call the add-in version, not a local version.
        strLibName = GetRunCmdAddInFullLibName
        strRunCmd = strLibName & ".APIAsync"

        ' Use Application.Run to call the add-in version
        If Not IsMissing(varArg2) Then
            APIAsync = Application.Run(strRunCmd, strCallbackInfo, strMethod, varArg1, varArg2)
        ElseIf Not IsMissing(varArg1) Then
            APIAsync = Application.Run(strRunCmd, strCallbackInfo, strMethod, varArg1)
        Else
            APIAsync = Application.Run(strRunCmd, strCallbackInfo, strMethod)
        End If
        GoTo CleanUp
    End If

    ' Determine if this is an async operation or should fall back to sync
    Select Case strMethod
        Case "Export", "FullExport", "ExportVBA", "Build", "BuildAs", "MergeBuild"
            ' These are async operations - spawn via timer with callback support

            ' Store callback info in registry for timer callback to retrieve
            SaveSetting PROJECT_NAME, "Timer", "CallbackInfo", strCallbackInfo
            MCPDebugLog "APIAsync: Method=" & strMethod & ", CallbackInfo length=" & Len(strCallbackInfo)
            MCPDebugLog "APIAsync: CallbackInfo=" & Left$(strCallbackInfo, 200)

            ' Determine timeout based on operation type
            Select Case strMethod
                Case "Export"
                    lngTimeoutMs = 300000  ' 5 minutes
                Case "FullExport"
                    lngTimeoutMs = 600000  ' 10 minutes (full export takes longer)
                Case "ExportVBA"
                    lngTimeoutMs = 120000  ' 2 minutes
                Case "Build", "BuildAs"
                    lngTimeoutMs = 600000  ' 10 minutes
                Case "MergeBuild"
                    lngTimeoutMs = 300000  ' 5 minutes
            End Select

            ' Use timer to spawn async operation
            ' The timer callback will read the callback info and start the operation
            If Not IsMissing(varArg2) Then
                SetTimer "APIAsyncOperation", strMethod, CStr(varArg1) & "|" & CStr(varArg2)
            ElseIf Not IsMissing(varArg1) Then
                SetTimer "APIAsyncOperation", strMethod, CStr(varArg1)
            Else
                SetTimer "APIAsyncOperation", strMethod, vbNullString
            End If

            ' Return async response immediately
            Set dResult = New Dictionary
            dResult.Add "async", True
            dResult.Add "timeout_ms", lngTimeoutMs
            APIAsync = modJsonConverter.ConvertToJson(dResult)

        Case Else
            ' Fall back to sync API for quick operations
            ' Call existing API function
            If Not IsMissing(varArg2) Then
                varResult = API(strMethod, varArg1, varArg2)
            ElseIf Not IsMissing(varArg1) Then
                varResult = API(strMethod, varArg1)
            Else
                varResult = API(strMethod)
            End If

            ' Return sync response
            Set dResult = New Dictionary
            dResult.Add "sync", True
            dResult.Add "result", varResult
            APIAsync = modJsonConverter.ConvertToJson(dResult)
    End Select

CleanUp:
    IsRunning = False
    RestoreErrorBreaks
    Exit Function

ErrHandler:
    ' An error occurred so we need to make it available for further attempts
    ' but do not handle the error.
    IsRunning = False
    RestoreErrorBreaks

    ' Re-throw
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function


'---------------------------------------------------------------------------------------
' Procedure : RefuseReentrantCall
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Build the message for a call that arrived while another is still running.
'           :
'           : This used to be a bare `Exit Function`, returning Empty -- indistinguishable
'           : from a method that legitimately returned nothing. The most common way to
'           : reach it makes that silence actively misleading: code submitted through
'           : vcs_run_vba is itself delivered via API, so anything it calls back into API
'           : is nested by construction, and every call coming back Empty reads as a
'           : broken add-in rather than a refused call.
'           :
'           : Returning a marked string rather than raising is deliberate, and was
'           : arrived at the hard way. An Err.Raise here does not reach the caller: an
'           : error raised inside a library database does not propagate across
'           : Application.Run into the calling project's handler, so even a caller with
'           : `On Error GoTo` active gets a modal "Run-time error" dialog that blocks
'           : Access until a human dismisses it. Since this guard only ever trips on a
'           : nested call -- precisely the case that crosses that boundary -- raising is
'           : guaranteed to hit it. A blocking dialog is worse for automation than the
'           : silence it replaced. Callers match API_REFUSED_PREFIX to tell a refusal
'           : from a real return value.
'---------------------------------------------------------------------------------------
'
Private Function RefuseReentrantCall(strMethod As String, strEntryPoint As String) As String
    RefuseReentrantCall = API_REFUSED_PREFIX & _
        "VCS " & strEntryPoint & " refused a re-entrant call to '" & strMethod & "': " & _
        "another API command is still running. Note that code executed through " & _
        "vcs_run_vba already runs inside an API call, so it cannot call back into the " & _
        "API; invoke API methods directly (vcs_call_vba) instead."
End Function


'---------------------------------------------------------------------------------------
' Procedure : MCPDebugLog
' Author    : Adam Waller
' Date      : 1/26/2026
' Purpose   : Write debug messages to a file for MCP callback troubleshooting.
'           : Writes to logs/MCP_Debug.log in the source folder.
'---------------------------------------------------------------------------------------
'
Public Sub MCPDebugLog(strMessage As String)

    ' No logging unless MCP is active
    If Not MCP.IsActive Then Exit Sub

    ' Initialize log path if needed
    If m_strMCPDebugLogPath = vbNullString Then
        m_strMCPDebugLogPath = Options.GetExportFolder & "logs\MCP_Debug.log"
    End If

    ' Append to log file
    AppendToFile Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & strMessage, m_strMCPDebugLogPath

End Sub
