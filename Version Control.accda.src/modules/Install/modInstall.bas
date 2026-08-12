Attribute VB_Name = "modInstall"
'---------------------------------------------------------------------------------------
' Module    : modInstall
' Author    : Adam Waller
' Date      : 2/4/2021
' Purpose   : This module contains the logic for installing/updating/removing/deploying
'           : the add-in.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Install")


' Registry hive
Private Enum eHive
    ehHKLM
    ehHKCU
End Enum

' Used to determine if Access is running as administrator. (Required for installing the add-in)
Private Declare PtrSafe Function IsUserAnAdmin Lib "shell32" () As Long

' Used by GetOtherAccessInstances to skip this process and other Windows sessions.
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Function ProcessIdToSessionId Lib "kernel32" ( _
    ByVal dwProcessId As Long, ByRef pSessionId As Long) As Long

' Used to describe the other Access instances that blocked a rebuild.
Private Declare PtrSafe Function EnumWindows Lib "user32" (ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal uCmd As Long) As LongPtr
Private Declare PtrSafe Function GetClassNameA Lib "user32" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function FindWindowExA Lib "user32" (ByVal hWndParent As LongPtr, ByVal hWndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As LongPtr, ByVal dwObjectID As Long, riid As Any, ppvObject As Object) As Long

Private Type udtIID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Const GW_OWNER As Long = 4
Private Const OBJID_NATIVEOM As Long = &HFFFFFFF0
Private Const S_OK As Long = 0

' EnumWindows callback state for a single PID inspection.
Private m_lngEnumTargetPid As Long
Private m_blnEnumHasVisible As Boolean
Private m_hwndEnumOMain As LongPtr

' What could be observed about another Access process. blnRespondedToAutomation
' records whether the native object model could be reached at all: nothing may act on
' the other two flags unless it is True, since an unreachable instance is
' indistinguishable from an idle one.
Private Type udtAccessInstance
    lngPid As Long
    blnRespondedToAutomation As Boolean
    blnHasVisibleWindow As Boolean
    blnDatabaseKnown As Boolean
    strDatabase As String
End Type

Private Const ModuleName As String = "modInstall"

' Used to add a trusted location for the add-in path (when necessary)
Private Const mcstrTrustedLocationName = PROJECT_NAME & " Version Control"

' Use a private type to manage install settings.
Public Type udtInstallSettings
    blnTrustAddInFolder As Boolean
    blnUseRibbonAddIn As Boolean
    blnUseCompiledAddIn As Boolean
    blnOpenAfterInstall As Boolean
    blnUseWorkerScript As Boolean
    strInstallFolder As String
    strSourcePath As String
    blnSettingsLoaded As Boolean
End Type
Private this As udtInstallSettings

Private m_blnInstallErrorTrappingActive As Boolean
Private m_intSavedInstallErrorTrapping As eVbeErrorTrapping


'---------------------------------------------------------------------------------------
' Procedure : EnterInstallErrorTrapping
' Author    : Adam Waller
' Date      : 7/17/2026
' Purpose   : Stage VBE error trapping for install/uninstall automation. Idempotent.
'---------------------------------------------------------------------------------------
'
Public Sub EnterInstallErrorTrapping()
    If Not m_blnInstallErrorTrappingActive Then
        m_intSavedInstallErrorTrapping = SaveUserErrorTrapping()
        ApplyUserErrorTrapping eetBreakInClassModule
        m_blnInstallErrorTrappingActive = True
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LeaveInstallErrorTrapping
' Author    : Adam Waller
' Date      : 7/17/2026
' Purpose   : Restore the user's VBE error trapping setting after install automation.
'---------------------------------------------------------------------------------------
'
Public Sub LeaveInstallErrorTrapping()
    If m_blnInstallErrorTrappingActive Then
        RestoreUserErrorTrapping m_intSavedInstallErrorTrapping
        m_blnInstallErrorTrappingActive = False
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ParseInstallCommand
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Normalize /cmd tokens for the installer. Returns "INSTALL",
'           : "INSTALL SILENT", or an empty string when the command is not an install.
'---------------------------------------------------------------------------------------
'
Public Function ParseInstallCommand(strCommand As String) As String

    Dim strNorm As String

    strNorm = UCase$(Trim$(strCommand))
    Do While InStr(strNorm, "  ") > 0
        strNorm = Replace(strNorm, "  ", " ")
    Loop

    Select Case strNorm
        Case "INSTALL", "INSTALL SILENT"
            ParseInstallCommand = strNorm
        Case Else
            ParseInstallCommand = vbNullString
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : AutoRun
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : This code runs when the add-in file is opened directly. It provides the
'           : user an easy way to update the add-in on their system.
'---------------------------------------------------------------------------------------
'
Public Function AutoRun() As Boolean

    Dim strInstallCmd As String

    ' Handle command-line install automation (/cmd INSTALL or /cmd INSTALL SILENT)
    strInstallCmd = ParseInstallCommand(Command$)
    If Len(strInstallCmd) > 0 Then
        If strInstallCmd = "INSTALL SILENT" Then
            Operation.Source = eosExternalAPI
            SetInteractionMode eimSilent
        End If
        VerifyResources
        GetInstallSettings
        ' Pass the saved worker script preference through rather than relying on the
        ' parameter default, so an unattended reinstall cannot silently re-enable a
        ' helper script the user deliberately turned off.
        InstallVCSAddin this.blnTrustAddInFolder, this.blnUseRibbonAddIn, _
            False, this.strInstallFolder, this.blnUseCompiledAddIn, _
            this.blnUseWorkerScript
        Exit Function
    End If

    ' A COM client is holding this instance and there is nobody to read a message box.
    ' Both branches below assume a person: one closes the instance out from under the
    ' client, the other strands it behind a form. That is what stops an agent from
    ' opening the add-in to run its own test suite, since those tests only run when the
    ' add-in is the current database.
    If OpenedByAutomation Then
        VerifyResources
        Exit Function
    End If

    ' See if the we are opening the file from the installed location.
    If CodeProject.FullName = GetInstalledAddInFileName Then

        ' Opening the file from add-in location, which would normally be unusual unless we are trying to remove
        ' legacy registry entries, or to trust the file after install.
        If IsUserAnAdmin = 1 Then RemoveLegacyInstall

        ' Adding a message box to here to autoclose the addin once the prompt is cleared.
        ' This handles the last step of the install for users that just installed the file.
        ' Since no code will run until the "Trust Document/Enable" is completed, this allows for the trust
        ' process to complete then close itself (if desired).

        ' For users that need to open the add-in file to trust it, show the confirmation
        ' message that the add-in has been installed successfully.
        MsgBox2 "Installation Complete!", _
            "You did it! Add-in version " & AppVersion & " is now installed.", _
            "Please reopen any instances of Microsoft Access before using the add-in." & vbCrLf & _
            "This instance of Microsoft Access will now close.", vbInformation
        DoCmd.Quit

    Else
        ' Could be running it from another location, such as after downloading
        ' an updated version of the addin, or building from source.
        VerifyResources

        ' Open installer form
        Form_frmVCSInstall.Visible = True
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : OpenedByAutomation
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : True when a COM client started this instance rather than a person.
'           : Returns False when the property cannot be read: suppressing the install
'           : UI for a real user is the worse of the two failures, so an unknown answer
'           : behaves as though someone is watching.
'---------------------------------------------------------------------------------------
'
Public Function OpenedByAutomation() As Boolean
    LogUnhandledErrors
    On Error Resume Next
    OpenedByAutomation = Not Application.UserControl
    If Err.Number <> 0 Then
        Err.Clear
        OpenedByAutomation = False
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : InstallVCSAddin
' Author    : Adam Waller
' Date      : 10/19/2020
' Purpose   : Installs/updates the add-in for the current user.
'           : Returns true if successful.
'---------------------------------------------------------------------------------------
'
Public Sub InstallVCSAddin(blnTrustFolder As Boolean, blnUseRibbon As Boolean, blnOpenAfterInstall As Boolean, strInstallFolder As String, _
            Optional ByVal blnCreateCompiledVersion As Boolean = False, _
            Optional ByVal blnUseWorkerScript As Boolean = True)

    Const OPEN_MODE_OPTION As String = "Default Open Mode for Databases"

    Dim strSource As String
    Dim strDest As String
    Dim strSilentFail As String

    EnterInstallErrorTrapping
    On Error GoTo CleanExit

    ' Verify the add-in file has the required name
    If StrComp(FSO.GetBaseName(CodeProject.Name), ADDIN_BASENAME, vbTextCompare) <> 0 Then
        MsgBox2 T("Unable to Install"), _
            T("The add-in file must be named ""{0}.accda"" to install correctly.", _
                var0:=ADDIN_BASENAME), _
            T("Please rename the file and try again."), vbExclamation
        GoTo CleanExit
    End If

    ' Load install settings from registry, then update with parameter values
    GetInstallSettings
    With this
        .blnUseRibbonAddIn = blnUseRibbon
        .blnUseCompiledAddIn = blnCreateCompiledVersion
        .blnOpenAfterInstall = blnOpenAfterInstall
        .blnTrustAddInFolder = blnTrustFolder
        .blnUseWorkerScript = blnUseWorkerScript
        If .strInstallFolder <> strInstallFolder Then
            ' Attempt to migrate any saved user settings files
            MigrateUserFiles strInstallFolder, GetFilePathsInFolder(.strInstallFolder)
            ' Update install folder to new path
            .strInstallFolder = strInstallFolder
        End If
   End With

    ' Save the updated settings to the registry.
    SaveInstallSettings

    ' Load some path values
    strSource = CodeProject.FullName
    strDest = GetAddInFileName
    VerifyPath strDest

    ' We can't replace a file with itself.  :-)
    If strSource = strDest Then
        MsgBox2 "Unable to Install", "You can't install the add-in over itself.", _
            "Please run from a different location to update.", , vbExclamation
        GoTo CleanExit
    End If

    ' Check default database open mode.
    If Application.GetOption(OPEN_MODE_OPTION) = 1 Then
        If Operation.InteractionMode = eimSilent Then
            ' Changing the option does not affect this already-open session, so the
            ' install still cannot continue. Record the change and ask the caller to retry.
            Application.SetOption OPEN_MODE_OPTION, 0
            Log.Add "Default Open Mode was Exclusive; changed to Shared. Retry the rebuild."
            strSilentFail = "Default Open Mode was Exclusive and has been changed to Shared. Retry the rebuild."
            GoTo CleanExit
        End If
        If MsgBox2("Default Open Mode set to Exclusive", _
            "The default open mode option for Microsoft Access is currently set to open databases in Exclusive mode by default. " & vbCrLf & _
            "This add-in needs to be opened in shared mode in order to install successfully.", _
            "Change the default open mode to 'Shared'?", vbYesNo + vbExclamation) = vbYes Then
            Application.SetOption OPEN_MODE_OPTION, 0
            MsgBox2 "Default Option Changed", _
                "Please restart Microsoft Access and run the install again.", , vbInformation
        End If
        GoTo CleanExit
    End If

    ' Run any applicable upgrades
    RunUpgrades

    ' Verify the trusted location
    If this.blnTrustAddInFolder Then VerifyTrustedLocation

    ' Copy the add-in file
    If Not UpdateAddInFile(blnCreateCompiledVersion) Then GoTo CleanExit

    ' Install the ribbon
    If this.blnUseRibbonAddIn Then
        ' Ensure that the ribbon is installed
        modCOMAddIn.VerifyComAddIn
    Else
        ' Remove if currently installed
        modCOMAddIn.UninstallComAddIn
    End If

    ' Remove the helper script when the user has turned it off, so a disabled feature does
    ' not leave its artifact in the add-in folder. Some endpoint protection objects to the
    ' file being there at all, not just to it running (#727). VerifyWorker writes it again
    ' on demand if the setting is turned back on.
    If Not this.blnUseWorkerScript Then Worker.RemoveWorkerScript

    ' Register the Menu controls
    RegisterMenuItem "&VCS Open", "=AddInMenuItemLaunch()"
    RegisterMenuItem "&VCS Options", "=AddInOptionsLaunch()"
    RegisterMenuItem "&VCS Export All Source", "=AddInMenuItemExport()"

    ' Update installed version number
    InstalledVersion = AppVersion

    ' Warn the user if ActiveX is disabled
    VerifyActivexNotDisabled

    ' Show install confirmation message
    MsgBox2 "Success!", "Version Control System add-in has been updated to " & AppVersion & ".", _
        "The installer will now close. Please restart any open instances" & vbCrLf & _
        "of Microsoft Access before using the add-in.", vbInformation

    ' Open add-in from installed location if required.
    If this.blnOpenAfterInstall Then OpenAddinFile GetAddInFileName, CodeProject.FullName

    ' Restore before quit so the user's setting is not left staged in the registry.
    LeaveInstallErrorTrapping

    RecordSilentInstallResult "complete"

    ' Close Access after installation is complete.
    DoCmd.Quit
    Exit Sub

CleanExit:
    LeaveInstallErrorTrapping
    If Operation.InteractionMode = eimSilent Then
        If Len(strSilentFail) = 0 Then
            If Err.Number <> 0 Then
                strSilentFail = Err.Description
            Else
                strSilentFail = "Install did not complete."
            End If
        End If
        RecordSilentInstallResult "install-failed", strSilentFail
    End If
    If Err.Number <> 0 Then
        MsgBox2 T("Unable to Install"), Err.Description, , vbExclamation
        Err.Clear
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : UninstallVCSAddin
' Author    : Adam Kauffman
' Date      : 5/27/2020
' Purpose   : Removes the add-in for the current user.
'           : Returns true if successful.
'---------------------------------------------------------------------------------------
'
Public Sub UninstallVCSAddin()

    Dim intResponse As VbMsgBoxResult
    Dim blnSaveSettings As Boolean
    Dim blnUseWorker As Boolean
    Dim strAddInFile As String
    Dim strAddInFolder As String

    EnterInstallErrorTrapping
    On Error GoTo CleanExit

    ' Read the install settings before the registry keys are removed below. Reading them
    ' afterwards would fall back to defaults, which is the wrong add-in path for anyone
    ' who installed to a custom folder, and would lose the user's worker preference.
    blnUseWorker = UseWorkerScript
    strAddInFile = GetInstalledAddInFileName
    strAddInFolder = GetInstallSettings.strInstallFolder

    ' Ask the user if they want to preserve their user settings.
    intResponse = MsgBox2("Save User Settings", "Would you like your user settings/options preserved?", _
        "Click YES to save these items so they can be used if you reinstall the add-in," & vbCrLf & _
        "Or click NO to remove all settings related to this add-on.", vbQuestion + vbYesNoCancel)

    ' Allow user to cancel if they are not sure how to answer the above prompt.
    If intResponse = vbCancel Then GoTo CleanExit

    ' Note if the user wants to save/migrate their existing settings.
    If intResponse = vbYes Then blnSaveSettings = True

    ' Close all database objects
    If IsLoaded(acForm, "frmVCSOptions") Then DoCmd.Close acForm, "frmVCSOptions"
    If IsLoaded(acForm, "frmVCSMain") Then DoCmd.Close acForm, "frmVCSMain"

    ' Remove the add-in Menu controls
    RemoveMenuItem "&VCS Open"
    RemoveMenuItem "&VCS Options"
    RemoveMenuItem "&VCS Export All Source"

    ' Remove any legacy menu items.
    RemoveMenuItem "&Version Control"
    RemoveMenuItem "&Version Control Options"
    RemoveMenuItem "&Export All Source"

    ' Remove registry entries
    LogUnhandledErrors
    On Error Resume Next
    If blnSaveSettings Then
        ' Delete keys that don't contain settings
        DeleteSetting PROJECT_NAME, "Build"
        DeleteSetting PROJECT_NAME, "Add-In"
        DeleteSetting PROJECT_NAME, "Timer"
    Else
        ' Remove entire application key
        DeleteSetting PROJECT_NAME
    End If

    ' Resume normal error handling
    If DebugMode(False) Then On Error GoTo CleanExit Else On Error Resume Next

    ' Remove trusted location added by this add-in. (if found)
    RemoveTrustedLocation

    ' Remove COM add-in
    modCOMAddIn.UninstallComAddIn

    ' Remove On Save hook
    'modExportOnSaveHook.Uninstall

    If blnUseWorker Then

        ' Notify the user of the completion of the uninstall process.
        MsgBox2 "Success!", "Version Control System has now been uninstalled.", _
            "Microsoft Access will be closed to remove the remaining files.", _
            vbInformation

        LeaveInstallErrorTrapping

        ' Use the worker script to actually remove the add-in files.
        ' (They cannot be removed when they are in use, such as when procesing the uninstall.)
        Worker.Run_UninstallAddin

    Else

        ' Without the helper script there is nothing that outlives this process to delete
        ' files it still holds open, so the last step becomes the user's. Access closes
        ' either way -- the files cannot be removed until it does. (#727)
        ' The script itself is only ever read by wscript, never held open by Access, so it
        ' can go now; NotifyManualAddInCleanup still lists it if this fails.
        Worker.RemoveWorkerScript
        NotifyManualAddInCleanup strAddInFile, strAddInFolder
        LeaveInstallErrorTrapping
        DoCmd.Quit

    End If
    Exit Sub

CleanExit:
    LeaveInstallErrorTrapping
    If Err.Number <> 0 Then
        MsgBox2 T("Unable to Uninstall"), Err.Description, , vbExclamation
        Err.Clear
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : NotifyManualAddInCleanup
' Author    : Adam Waller
' Date      : 8/7/2026
' Purpose   : Tell the user which files are left behind after an uninstall that could not
'           : use the helper script, and open the folder so they can remove them once
'           : Access has closed.
'           :
'           : Naming the files matters here: the add-in database and its lock file are the
'           : ones the user in #727 was left with and had no way to identify.
'---------------------------------------------------------------------------------------
'
Private Sub NotifyManualAddInCleanup(strAddInFile As String, strFolder As String)

    Dim strFiles As String

    ' List what is actually left in the folder.
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add FSO.GetFileName(strAddInFile)
        .Add FSO.GetBaseName(strAddInFile) & ".laccdb"
        If FSO.FileExists(BuildPath2(strFolder, "Worker.vbs")) Then .Add "Worker.vbs"
        strFiles = .GetStr
    End With

    MsgBox2 T("Almost Finished"), _
        T("Version Control System has been uninstalled, but these files cannot be " & _
        "removed while Microsoft Access has them open:") & vbCrLf & vbCrLf & strFiles, _
        T("Microsoft Access will now close. Please delete them from this folder:") & _
        vbCrLf & strFolder, vbInformation

    ' Open the folder so the files are in front of the user when Access closes. Not worth
    ' failing the uninstall over if the shell declines to open it.
    LogUnhandledErrors
    On Error Resume Next
    Application.FollowHyperlink strFolder
    If Err Then Err.Clear
    On Error GoTo 0

End Sub


'---------------------------------------------------------------------------------------
' Procedure : UpdateAddInFile
' Author    : Adam Waller
' Date      : 5/22/2023
' Purpose   : Update the add-in database file. Return true if successful.
'---------------------------------------------------------------------------------------
'
Private Function UpdateAddInFile(ByVal blnCreateCompiledVersion As Boolean) As Boolean

    Dim strAddInFile As String
    Dim strCompiledFile As String

    ' Build file paths before entering the error-handled block so that
    ' no calls to the FSO() getter (which contains LogUnhandledErrors)
    ' can intercept a pending error before we check for it below.
    strAddInFile = GetAddInFileName
    strCompiledFile = GetAddInFileName(True)

    ' Make sure the destination folder exists
    VerifyPath strAddInFile

    ' Update the file
    LogUnhandledErrors
    On Error GoTo UpdateError
    If FSO.FileExists(strAddInFile) Then DeleteFile strAddInFile, True

    If blnCreateCompiledVersion Then
        ' Remove any existing uncompiled version
        ' (Very important, since we are invoking the add-in file directly without the extension.
        ' if both files exist, the .accda file will be opened instead of the compiled one.)
        If FSO.FileExists(strAddInFile) Then DeleteFile strAddInFile
        ' Now we can generate the compiled version as a *.accde
        CreateAccde CodeProject.FullName, strCompiledFile
    Else
        FSO.CopyFile CodeProject.FullName, strAddInFile, True
    End If

    ' Remove any existing alternate version
    If Not blnCreateCompiledVersion Then
        If FSO.FileExists(strCompiledFile) Then DeleteFile strCompiledFile
    End If
    ' Copied file with no errors.
    UpdateAddInFile = True
    Exit Function

UpdateError:
    MsgBox2 "Unable to Update File", _
        "Encountered error " & Err.Number & ": " & Err.Description & " when copying file.", _
        "Is the Version Control Add-in loaded in another instance of Microsoft Access?" & vbCrLf & _
        "Please check to be sure that the following file is not in use:" & _
        vbCrLf & strAddInFile, vbExclamation

End Function


'---------------------------------------------------------------------------------------
' Procedure : CreateAccde
' Author    : Josef Poetzl
' Date      : 2/21/2025
' Purpose   : Create a compiled Access file
'---------------------------------------------------------------------------------------
'
Private Sub CreateAccde(ByVal strSourceFilePath As String, ByVal strDestFilePath As String)

    Const acSysCmdCompile As Long = 603 ' Added in later versions of Access
    Dim strFileToCompile As String
    Dim objAccess As Access.Application
    Dim intSavedErrorTrapping As eVbeErrorTrapping
    Dim lngErr As Long
    Dim strErrDesc As String

    strFileToCompile = strDestFilePath & ".accdb"
    FSO.CopyFile strSourceFilePath, strFileToCompile, True

    ' use new Access instance to create accde
    Set objAccess = New Access.Application
    intSavedErrorTrapping = SaveUserErrorTrappingOnApp(objAccess)
    ApplyUserErrorTrappingOnApp objAccess, eetBreakInClassModule
    On Error GoTo CreateAccdeErr
    objAccess.Visible = True
    objAccess.SysCmd acSysCmdCompile, strFileToCompile, strDestFilePath
    GoTo CreateAccdeExit

CreateAccdeErr:
    lngErr = Err.Number
    strErrDesc = Err.Description

CreateAccdeExit:
    On Error Resume Next
    RestoreUserErrorTrappingOnApp objAccess, intSavedErrorTrapping
    If Not objAccess Is Nothing Then objAccess.Quit
    Set objAccess = Nothing
    FSO.DeleteFile strFileToCompile, True
    If lngErr <> 0 Then Err.Raise lngErr, , strErrDesc

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MigrateUserFiles
' Author    : Adam Waller
' Date      : 5/22/2023
' Purpose   : Migrate a file, unless a destination file exists that is different from
'           : the source file.
'---------------------------------------------------------------------------------------
'
Private Sub MigrateUserFiles(strToFolder As String, colNames As Dictionary)

    Dim varKey As Variant
    Dim strFile As String
    Dim strSource As String
    Dim strDest As String

    ' Loop through file names
    For Each varKey In colNames.Keys
        strSource = varKey
        strFile = FSO.GetFileName(strSource)
        strDest = BuildPath2(strToFolder, strFile)
        Select Case True
            ' Define exceptions to skip
            Case strFile Like PROJECT_NAME & ".*accda"  ' Add-in or lock file
            Case strFile Like "*.dll"   ' COM dlls
            Case strFile Like "*.vbs"   ' Worker script
            Case Else
                ' Migrate other files
                If FSO.FileExists(strSource) Then
                    If FSO.FileExists(strDest) Then
                        ' Check hash of file content
                        If GetFileHash(strSource) = GetFileHash(strDest) Then
                            ' File is identical in content. Remove source file.
                            DeleteFile strSource
                        Else
                            ' Leave existing file if they don't match.
                        End If
                    Else
                        ' If destination file does not exist, move from source.
                        FSO.MoveFile strSource, strDest
                    End If
                End If
        End Select
    Next varKey

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAddinFileName
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : This is where the add-in would be installed.
'---------------------------------------------------------------------------------------
'
Public Function GetAddInFileName(Optional blnAsMde As Boolean = False) As String
    GetAddInFileName = FSO.BuildPath(GetInstallSettings.strInstallFolder, _
        ADDIN_BASENAME & IIf(blnAsMde, ".accde", ".accda"))
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetInstalledAddInFileName
' Author    : Adam Waller
' Date      : 3/27/2026
' Purpose   : Returns the full path to the installed add-in file, using the correct
'           : extension (.accda or .accde) based on the persisted install setting.
'---------------------------------------------------------------------------------------
'
Public Function GetInstalledAddInFileName() As String
    GetInstalledAddInFileName = GetAddInFileName(GetInstallSettings.blnUseCompiledAddIn)
End Function


'---------------------------------------------------------------------------------------
' Procedure : UseWorkerScript
' Author    : Adam Waller
' Date      : 8/7/2026
' Purpose   : Returns true when the add-in is allowed to extract and run the helper
'           : VBScript (`Worker.vbs`). Some managed environments have endpoint
'           : protection that blocks Access from launching an extracted script, which
'           : prevented those users from running v5 at all. (See issue #727.)
'           : This is a per-user install setting rather than a project option, because
'           : uninstall and add-in rebuild are not scoped to any project, and the reason
'           : to disable it belongs to the machine rather than the source tree.
'           : Callers do not normally need this: `clsWorker.CallWorker` checks it, so a
'           : disabled worker turns every job into a no-op. Read it directly only where
'           : a different code path is needed rather than a skipped one.
'---------------------------------------------------------------------------------------
'
Public Function UseWorkerScript() As Boolean
    UseWorkerScript = GetInstallSettings.blnUseWorkerScript
End Function


'---------------------------------------------------------------------------------------
' Procedure : DefaultAddInFolderPath
' Author    : Adam Waller
' Date      : 5/22/2023
' Purpose   : Returns the default installation folder path.
'---------------------------------------------------------------------------------------
'
Private Function DefaultAddInFolderPath() As String
    DefaultAddInFolderPath = BuildPath2(Environ$("AppData"), PROJECT_NAME)
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddinLoaded
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : Returns true if the VCS add-in is currently loaded as a VBE Project.
'---------------------------------------------------------------------------------------
'
Public Function AddinLoaded() As Boolean
    AddinLoaded = Not GetAddInProject Is Nothing
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAddinRegPath
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Return the registry path to the addin menu items
'---------------------------------------------------------------------------------------
'
Private Function GetAddinRegPath(Optional Hive As eHive = ehHKCU) As String

    Dim strHive As String

    Select Case Hive
        Case ehHKCU: strHive = "HKCU\"
        Case ehHKLM: strHive = "HKLM\"
    End Select

    GetAddinRegPath = strHive & "SOFTWARE\Microsoft\Office\" & _
            Application.Version & "\Access\Menu Add-Ins\"

End Function


'---------------------------------------------------------------------------------------
' Procedure : RegisterMenuItem
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Add the menu item through the registry (Normally HKCU hive)
'---------------------------------------------------------------------------------------
'
Private Sub RegisterMenuItem(ByVal strName As String, Optional ByVal strFunction As String = "=LaunchMe()")

    Dim strPath As String

    ' We need to create/update three registry keys for each item.
    strPath = GetAddinRegPath & strName & "\"
    With New IWshRuntimeLibrary.WshShell
        .RegWrite strPath & "Expression", strFunction, "REG_SZ"
        .RegWrite strPath & "Library", GetInstalledAddInFileName, "REG_SZ"
        .RegWrite strPath & "Version", 3, "REG_DWORD"
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveMenuItem
' Author    : Adam Kauffman
' Date      : 5/27/2020
' Purpose   : Remove the menu item through the registry
'---------------------------------------------------------------------------------------
'
Private Sub RemoveMenuItem(ByVal strName As String, Optional Hive As eHive = ehHKCU)

    Dim strPath As String

    ' We need to remove three registry keys for each item.
    strPath = GetAddinRegPath(Hive) & strName & "\"
    With New IWshRuntimeLibrary.WshShell
        ' Just in case someone changed some of the keys...
        LogUnhandledErrors
        On Error Resume Next
        .RegDelete strPath & "Expression"
        .RegDelete strPath & "Library"
        .RegDelete strPath & "Version"
        .RegDelete strPath
        On Error GoTo 0
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RelaunchAsAdmin
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Launch the addin file with admin privileges so the user can uninstall it.
'---------------------------------------------------------------------------------------
'
Private Sub RelaunchAsAdmin()
    ShellEx FSO.BuildPath(SysCmd(acSysCmdAccessDir), "msaccess.exe"), """" & GetInstalledAddInFileName & """", "runas"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Deploy
' Author    : Adam Waller
' Date      : 4/21/2020
' Purpose   : Increments the build version and updates the project description.
'           : This can be run from the debug window when making updates to the project.
'           : (More significant updates to the version number can be made using the
'           :  `AppVersion` property defined below.)
'---------------------------------------------------------------------------------------
'
Public Sub Deploy(Optional ReleaseType As eReleaseType = Same_Version)

    Const cstrSpacer As String = "--------------------------------------------------------------"

    Dim strBinaryFile As String

    If Not IsCompiled Then
        MsgBox2 "Please Compile and Save Project", _
            "The project needs to be compiled and saved before deploying.", _
            "I would do this for you, but it seems to cause memory heap corruption" & vbCrLf & _
            "when this is run via VBA code during the deployment process." & vbCrLf & _
            "(This can be fixed by rebuilding from source.)", vbInformation
        Exit Sub
        ' Save all code modules
        'DoCmd.RunCommand acCmdCompileAndSaveAllModules
    End If

    If AddinLoaded Then
        MsgBox2 "Add-in Currently Loaded", _
            "The add-in file cannot be updated when it is currently in use.", _
            "Please close Microsoft Access and open this file again to deploy.", vbExclamation
        Exit Sub
    End If

    ' Make sure we don't run ths function while it is loaded in another project.
    If CodeProject.FullName <> CurrentProject.FullName Then
        Debug.Print "This can only be run from a top-level project."
        Debug.Print "Please open " & CodeProject.FullName & " and try again."
        Exit Sub
    End If

    ' Increment build number
    IncrementAppVersion ReleaseType

    ' List project and new build number
    Debug.Print cstrSpacer

    ' Update project description and save
    CurrentVBProject.Description = "Version " & AppVersion & " deployed on " & Date
    DoCmd.RunCommand acCmdCompileAndSaveAllModules

    ' Save copy to zip folder
    strBinaryFile = FSO.BuildPath(CodeProject.Path, "Version_Control_v" & AppVersion & ".zip")
    If FSO.FileExists(strBinaryFile) Then DeleteFile strBinaryFile, True
    CreateZipFile strBinaryFile
    CopyFileToZip CodeProject.FullName, strBinaryFile

    ' Deploy latest version on this machine
    If Not UpdateAddInFile(False) Then Exit Sub

    ' Use the newly installed add-in to Export the project to version control.
    modAPI.HandleRibbonCommand "btnExport"

    ' Finish with success message if the latest version was installed.
    Debug.Print "Version " & AppVersion & " installed."

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunUpgrades
' Author    : Adam Waller
' Date      : 5/27/2020
' Purpose   : Process upgrade transitions and remove legacy components
'---------------------------------------------------------------------------------------
'
Private Sub RunUpgrades()

    Dim strName As String
    Dim strOldPath As String
    Dim strNewPath As String
    Dim strTest As String
    Dim objShell As IWshRuntimeLibrary.WshShell

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Legacy HKLM install
    If InstalledVersion < "3.2.0" Then
        ' Check for installation in HKLM hive.
        strOldPath = GetAddinRegPath(ehHKLM) & "&Version Control\Library"
        Set objShell = New IWshRuntimeLibrary.WshShell
        LogUnhandledErrors
        On Error Resume Next
        strTest = objShell.RegRead(strOldPath)
        If Err Then Err.Clear
        On Error GoTo 0
        If strTest <> vbNullString Then
            If MsgBox2("Remove Legacy Version?", "Way back in the old days, this install required admin rights " & _
                "and added some keys to the HKLM registry. We don't need those anymore " & _
                "because the add-in is now installed for the current user with no special " & _
                "privileges required." _
                , "Can we go ahead and clean those up now? (Requires admin to remove the registry keys.)" _
                , vbQuestion + vbYesNo) = vbYes Then
                RelaunchAsAdmin
            End If
        End If
    End If

    ' Install in Microsoft\AddIns\ folder
    If InstalledVersion < "3.3.0" Then

        ' Check for install in AddIns folder (before we used the dedicated install folder)
        strOldPath = BuildPath2(Environ$("AppData"), "Microsoft", "AddIns", ADDIN_BASENAME & ".accda")

        ' Remove add-in from legacy location
        If FSO.FileExists(strOldPath) Then DeleteFile strOldPath

        ' Migrate settings json file to new location
        strOldPath = Replace(strOldPath, ".accda", ".json", , , vbTextCompare)
        If FSO.FileExists(strOldPath) Then
            ' Check for settings file in new location
            strNewPath = Replace(GetAddInFileName, ".accda", ".json", , , vbTextCompare)
            If FSO.FileExists(strNewPath) Then
                ' Leave new settings file, and delete old one.
                DeleteFile strOldPath
            Else
                ' Move settings to new location
                VerifyPath strNewPath
                FSO.MoveFile strOldPath, strNewPath
            End If
        End If

        ' Remove any Legacy Menu controls
        RemoveMenuItem "&Version Control"
        RemoveMenuItem "&Version Control Options"
        RemoveMenuItem "&Export All Source"

        ' Remove custom trusted location for Office AddIns folder.
        strName = "Office Add-ins"
        If HasTrustedLocationKey(strName) Then RemoveTrustedLocation strName
    End If

    ' Use standardized options folder (5/7/2021)
    strOldPath = FSO.BuildPath(CodeProject.Path, ADDIN_BASENAME) & ".json"
    strNewPath = FSO.BuildPath(CodeProject.Path, "vcs-options.json")
    If FSO.FileExists(strOldPath) Then
        If FSO.FileExists(strNewPath) Then
            ' Remove leftover legacy file
            DeleteFile strOldPath
        Else
            ' Rename to new name
            Name strOldPath As strNewPath
        End If
    End If

    ' Handle any uncaught errors
    CatchAny eelError, "Running upgrades before install", ModuleName & ".RunUpgrades"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveLegacyInstall
' Author    : Adam Waller
' Date      : 5/27/2020
' Purpose   : Remove the installation that required admin rights in favor of the
'           : per-user installation method.
'---------------------------------------------------------------------------------------
'
Private Sub RemoveLegacyInstall()

    ' These registry keys require admin access to remove
    RemoveMenuItem "&Version Control", ehHKLM
    RemoveMenuItem "&Export All Source", ehHKLM

    MsgBox2 "Legacy Items Removed", "Thanks for getting those cleaned up!" _
        , "Microsoft Access will now close so you can continue.", vbInformation
    DoCmd.Quit

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetInstallSettings
' Author    : Adam Waller
' Date      : 5/22/2023
' Purpose   : Return the install settings.
'---------------------------------------------------------------------------------------
'
Public Function GetInstallSettings(Optional blnUseCache As Boolean = True) As udtInstallSettings

    ' Load install settings from registry
    With this
        If Not (.blnSettingsLoaded And blnUseCache) Then
            .blnTrustAddInFolder = GetSetting(PROJECT_NAME, "Install", "Trust Folder", CInt(True))
            .blnUseRibbonAddIn = GetSetting(PROJECT_NAME, "Install", "Use Ribbon", True)
            .blnUseCompiledAddIn = GetSetting(PROJECT_NAME, "Install", "Compile accde", False)
            .blnOpenAfterInstall = GetSetting(PROJECT_NAME, "Install", "Open File", CInt(False))
            .blnUseWorkerScript = GetSetting(PROJECT_NAME, "Install", "Use Worker Script", CInt(True))
            .strInstallFolder = GetSetting(PROJECT_NAME, "Install", "Install Folder", DefaultAddInFolderPath)
            .strSourcePath = GetSetting(PROJECT_NAME, "Install", "Source Path", vbNullString)
            .blnSettingsLoaded = True
        End If
    End With
    GetInstallSettings = this

End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveInstallSettings
' Author    : Adam Waller
' Date      : 5/22/2023
' Purpose   : Saves current install settings to the registry.
'---------------------------------------------------------------------------------------
'
Public Function SaveInstallSettings()
    With this
        ' Basic settings
        SaveSetting PROJECT_NAME, "Install", "Trust Folder", CInt(.blnTrustAddInFolder)
        SaveSetting PROJECT_NAME, "Install", "Use Ribbon", CInt(.blnUseRibbonAddIn)
        SaveSetting PROJECT_NAME, "Install", "Compile accde", CInt(.blnUseCompiledAddIn)
        SaveSetting PROJECT_NAME, "Install", "Open File", CInt(.blnOpenAfterInstall)
        SaveSetting PROJECT_NAME, "Install", "Use Worker Script", CInt(.blnUseWorkerScript)
        If Len(.strSourcePath) Then
            SaveSetting PROJECT_NAME, "Install", "Source Path", .strSourcePath
        End If
        ' Special handling
        If .strInstallFolder = DefaultAddInFolderPath Then
            ' This value should only be saved if using a non-standard path.
            If GetSetting(PROJECT_NAME, "Install", "Install Folder") <> vbNullString Then
                ' Remove custom folder path setting when it matches the default.
                DeleteSetting PROJECT_NAME, "Install", "Install Folder"
            End If
        Else
            ' Save the custom path
            SaveSetting PROJECT_NAME, "Install", "Install Folder", .strInstallFolder
        End If
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : VerifyTrustedLocation
' Author    : Adam Waller
' Date      : 1/12/2021
' Purpose   : The location of the add-in must be trusted, or the user will be unable
'           : to run the add-in. This function ensures that the path has been added
'           : as a trusted location after confirming this with the user. If the user
'           : declines to add as a trusted location, it warns them that the add-in may
'           : not function correctly.
'---------------------------------------------------------------------------------------
'
Private Function VerifyTrustedLocation() As Boolean

    Dim strPath As String
    Dim strTrusted As String

    ' Get registry path for trusted locations
    strPath = GetTrustedLocationRegPath
    strTrusted = FSO.GetParentFolderName(GetAddInFileName) & PathSep

    ' Use Windows Scripting Shell to read/write to registry
    With New IWshRuntimeLibrary.WshShell

        ' Check for existing value
        If HasTrustedLocationKey Then

            ' Found trusted location with this name.
            VerifyTrustedLocation = True

        Else
            ' Get permission from user to add trusted location
            If MsgBox2("Add Trusted Location?", _
                "To function correctly, this add-in needs to be ""trusted"" by Microsoft Access." & vbCrLf & _
                "Typically this is accomplished by adding the add-in folder as a trusted location" & vbCrLf & _
                "in your security settings. More information is available on the GitHub wiki for" & vbCrLf & _
                "this add-in project.", _
                "<<PLEASE CONFIRM>> Add the following path as a trusted location?" & vbCrLf & vbCrLf & strTrusted _
                , vbQuestion + vbOKCancel + vbDefaultButton2) = vbOK Then

                ' Add trusted location
                .RegWrite strPath & "Path", strTrusted
                .RegWrite strPath & "Date", Now()
                .RegWrite strPath & "Description", mcstrTrustedLocationName
                .RegWrite strPath & "AllowSubfolders", 0, "REG_DWORD"

                ' Verify it was actually set.
                If HasTrustedLocationKey Then
                    VerifyTrustedLocation = True
                Else
                    ' Could not find registry entry.
                    MsgBox2 "Hmm... Something didn't work", _
                        "The new trusted location entry was not found in the registry.", _
                        "Please open an issue on GitHub if the issue persists.", vbExclamation
                End If

            Else
                MsgBox2 "Location NOT Added", _
                    "No problem. You can always run the installer again" & vbCrLf & _
                    "if you change your mind.", _
                    "Note that the add-in may not function correctly.", vbInformation
            End If
        End If
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveTrustedLocation
' Author    : Adam Waller
' Date      : 1/12/2021
' Purpose   : Remove trusted location entry.
'---------------------------------------------------------------------------------------
'
Public Sub RemoveTrustedLocation(Optional strName As String)

    Dim strPath As String

    ' Get registry path for trusted locations
    strPath = GetTrustedLocationRegPath(strName)

    With New IWshRuntimeLibrary.WshShell
        LogUnhandledErrors
        On Error Resume Next
        .RegDelete strPath & "Path"
        .RegDelete strPath & "Date"
        .RegDelete strPath & "Description"
        .RegDelete strPath & "AllowSubfolders"
        .RegDelete strPath
        On Error GoTo 0
    End With

    ' Make sure it was removed
    If HasTrustedLocationKey Then
        MsgBox2 "Error Removing Trusted Location", _
            "You may need to manually remove the trusted location" & vbCrLf & _
            "in the Microsoft Access Security settings.", _
            "Please open an issue on GitHub if the issue persists.", vbExclamation
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetTrustedLocationRegPath
' Author    : Adam Waller
' Date      : 1/12/2021
' Purpose   : Return the trusted location registry path. (Added to trusted locations)
'---------------------------------------------------------------------------------------
'
'
Private Function GetTrustedLocationRegPath(Optional ByVal strName As String) As String

    ' If no (other) name was specified, default to the standard one.
    If strName = vbNullString Then strName = mcstrTrustedLocationName

    ' Return the full registry path to the trusted location
    GetTrustedLocationRegPath = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & _
        Application.Version & "\Access\Security\Trusted Locations\" & strName & "\"

End Function


'---------------------------------------------------------------------------------------
' Procedure : HasTrustedLocationKey
' Author    : Adam Waller
' Date      : 1/13/2021
' Purpose   : Returns true if we find the trusted location added by this add-in.
'---------------------------------------------------------------------------------------
'
Public Function HasTrustedLocationKey(Optional strName As String) As Boolean
    With New IWshRuntimeLibrary.WshShell
        LogUnhandledErrors
        On Error Resume Next
        HasTrustedLocationKey = Nz(.RegRead(GetTrustedLocationRegPath(strName) & "Path")) <> vbNullString
        If Err Then Err.Clear
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetOtherAccessInstances
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Count MSACCESS.EXE processes in this Windows session other than the
'           : current process. Returns -1 if the process query itself failed (fail-safe:
'           : callers must refuse). strDetail describes each instance: its open
'           : database, whether it has a visible window, and whether its object model
'           : could be reached. Never quits or terminates anything.
'---------------------------------------------------------------------------------------
'
Public Function GetOtherAccessInstances(ByRef strDetail As String) As Long

    Dim dProcs As Dictionary
    Dim strError As String
    Dim varPid As Variant
    Dim udtInfo As udtAccessInstance
    Dim cDetail As clsConcat

    strDetail = vbNullString
    GetOtherAccessInstances = -1

    If Not CollectOtherAccessProcesses(dProcs, strError) Then
        strDetail = strError
        Exit Function
    End If

    Set cDetail = New clsConcat
    cDetail.AppendOnAdd = vbCrLf
    For Each varPid In dProcs.Keys
        udtInfo = ClassifyAccessInstance(CLng(varPid), Nz(dProcs(varPid), vbNullString))
        cDetail.Add DescribeAccessInstance(udtInfo)
    Next varPid

    strDetail = cDetail.GetStr
    GetOtherAccessInstances = dProcs.Count

End Function


'---------------------------------------------------------------------------------------
' Procedure : CollectOtherAccessProcesses
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Dictionary of other MSACCESS.EXE PIDs in this session mapped to their
'           : command lines. Returns False if the process query failed.
'---------------------------------------------------------------------------------------
'
Private Function CollectOtherAccessProcesses(ByRef dProcs As Dictionary, _
    ByRef strError As String) As Boolean

    Dim objWMI As Object
    Dim colProcs As Object
    Dim objProc As Object
    Dim lngThisPid As Long
    Dim lngThisSession As Long
    Dim lngPid As Long
    Dim lngSession As Long
    Dim strCmd As String

    strError = vbNullString
    Set dProcs = New Dictionary

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    lngThisPid = GetCurrentProcessId
    If ProcessIdToSessionId(lngThisPid, lngThisSession) = 0 Then
        strError = "Could not determine the current Windows session."
        CatchAny eelError, strError, ModuleName & ".CollectOtherAccessProcesses", True, True
        Exit Function
    End If

    Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    If objWMI Is Nothing Or Err.Number <> 0 Then
        strError = "Process query failed: " & Err.Description
        CatchAny eelError, strError, ModuleName & ".CollectOtherAccessProcesses", True, True
        Exit Function
    End If

    Set colProcs = objWMI.ExecQuery("SELECT ProcessId, CommandLine, SessionId FROM Win32_Process WHERE Name = 'MSACCESS.EXE'")
    If colProcs Is Nothing Or Err.Number <> 0 Then
        strError = "Process query failed: " & Err.Description
        CatchAny eelError, strError, ModuleName & ".CollectOtherAccessProcesses", True, True
        Exit Function
    End If

    For Each objProc In colProcs
        lngPid = 0
        lngSession = -1
        strCmd = vbNullString
        On Error Resume Next
        lngPid = CLng(objProc.ProcessId)
        lngSession = CLng(objProc.SessionId)
        If Not IsNull(objProc.CommandLine) Then strCmd = CStr(objProc.CommandLine)
        If Err Then Err.Clear
        If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

        If lngPid <> lngThisPid And lngSession = lngThisSession Then
            If Not dProcs.Exists(CStr(lngPid)) Then dProcs.Add CStr(lngPid), strCmd
        End If
    Next objProc

    If Err.Number <> 0 Then
        strError = "Process query failed: " & Err.Description
        CatchAny eelError, strError, ModuleName & ".CollectOtherAccessProcesses", True, True
        Exit Function
    End If

    CollectOtherAccessProcesses = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : ClassifyAccessInstance
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Gather what can be observed about another Access process without changing
'           : it. Command-line paths alone do not settle whether a database is open,
'           : because COM OpenCurrentDatabase does not appear there, so the native
'           : object model is also probed. A busy instance rejects automation calls and
'           : therefore looks the same as an idle one; blnRespondedToAutomation records
'           : which of the two it was, and callers must not treat an unreachable
'           : instance as idle.
'---------------------------------------------------------------------------------------
'
Private Function ClassifyAccessInstance(lngPid As Long, strCmd As String) As udtAccessInstance

    Dim objApp As Object
    Dim udtInfo As udtAccessInstance
    Dim strDb As String
    Dim strVersion As String

    udtInfo.lngPid = lngPid
    udtInfo.strDatabase = ExtractOpenDatabaseFromCommandLine(strCmd)

    InspectAccessWindows lngPid
    udtInfo.blnHasVisibleWindow = m_blnEnumHasVisible

    Set objApp = AccessAppFromHwnd(m_hwndEnumOMain)
    If objApp Is Nothing Then
        ClassifyAccessInstance = udtInfo
        Exit Function
    End If

    LogUnhandledErrors
    On Error Resume Next
    Err.Clear

    ' Version answers on any responsive instance, database or not, so it separates
    ' "told us it has nothing open" from "never answered".
    strVersion = objApp.Version
    udtInfo.blnRespondedToAutomation = (Err.Number = 0) And (Len(strVersion) > 0)
    Err.Clear

    If udtInfo.blnRespondedToAutomation Then
        strDb = objApp.CurrentProject.FullName
        If Err.Number = 0 Then
            udtInfo.blnDatabaseKnown = True
            If Len(strDb) > 0 Then udtInfo.strDatabase = strDb
        End If
        Err.Clear
        If objApp.Visible Then udtInfo.blnHasVisibleWindow = True
        Err.Clear
    End If

    Set objApp = Nothing
    ClassifyAccessInstance = udtInfo

End Function


'---------------------------------------------------------------------------------------
' Procedure : DescribeAccessInstance
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : One line per instance for the refusal message, carrying the evidence a
'           : user (or agent) needs to decide what to close by hand.
'---------------------------------------------------------------------------------------
'
Private Function DescribeAccessInstance(udtInfo As udtAccessInstance) As String

    Dim strDb As String
    Dim strState As String

    If Len(udtInfo.strDatabase) > 0 Then
        strDb = udtInfo.strDatabase
    ElseIf udtInfo.blnDatabaseKnown Then
        strDb = T("no database open")
    Else
        strDb = T("open database unknown")
    End If

    If udtInfo.blnHasVisibleWindow Then strState = T("visible") Else strState = T("hidden")
    If udtInfo.blnRespondedToAutomation Then
        strState = strState & ", " & T("responded to automation")
    Else
        strState = strState & ", " & T("did not respond to automation")
    End If

    DescribeAccessInstance = "PID " & udtInfo.lngPid & ": " & strDb & " (" & strState & ")"

End Function


'---------------------------------------------------------------------------------------
' Procedure : InspectAccessWindows
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Record whether PID has a visible top-level window and capture its OMain
'           : hwnd for AccessibleObjectFromWindow.
'---------------------------------------------------------------------------------------
'
Private Sub InspectAccessWindows(lngPid As Long)
    m_lngEnumTargetPid = lngPid
    m_blnEnumHasVisible = False
    m_hwndEnumOMain = 0
    EnumWindows AddressOf EnumAccessWindowsCallback, 0
End Sub


'---------------------------------------------------------------------------------------
' Procedure : EnumAccessWindowsCallback
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : EnumWindows callback. Must live in a standard module for AddressOf.
'           : An error raised here would unwind through Win32 stack frames and take
'           : Access down with it, so this never lets one escape and never reports
'           : failure to the caller.
'---------------------------------------------------------------------------------------
'
Private Function EnumAccessWindowsCallback(ByVal hwnd As LongPtr, ByVal lParam As LongPtr) As Long

    Dim lngPid As Long
    Dim strClass As String
    Dim lngLen As Long

    EnumAccessWindowsCallback = 1
    On Error Resume Next

    GetWindowThreadProcessId hwnd, lngPid
    If lngPid = m_lngEnumTargetPid Then
        If GetWindow(hwnd, GW_OWNER) = 0 Then
            If IsWindowVisible(hwnd) <> 0 Then m_blnEnumHasVisible = True
        End If
        strClass = String$(64, vbNullChar)
        lngLen = GetClassNameA(hwnd, strClass, 64)
        If lngLen > 0 Then
            If StrComp(Left$(strClass, lngLen), "OMain", vbTextCompare) = 0 Then
                m_hwndEnumOMain = hwnd
            End If
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : AccessAppFromHwnd
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Native Access.Application from an OMain hwnd, or Nothing.
'---------------------------------------------------------------------------------------
'
Private Function AccessAppFromHwnd(ByVal hwnd As LongPtr) As Object

    Dim iid As udtIID
    Dim obj As Object
    Dim hwndChild As LongPtr

    If hwnd = 0 Then Exit Function

    ' A process that exits between enumeration and this call fails the marshalled
    ' call rather than the guard.
    LogUnhandledErrors
    On Error Resume Next

    iid.Data1 = &H20400
    iid.Data4(0) = &HC0
    iid.Data4(7) = &H46

    If AccessibleObjectFromWindow(hwnd, OBJID_NATIVEOM, iid, obj) = S_OK Then
        Set AccessAppFromHwnd = ApplicationFromComObject(obj)
        If Not AccessAppFromHwnd Is Nothing Then Exit Function
    End If

    hwndChild = FindWindowExA(hwnd, 0, vbNullString, vbNullString)
    Do While hwndChild <> 0
        Set obj = Nothing
        If AccessibleObjectFromWindow(hwndChild, OBJID_NATIVEOM, iid, obj) = S_OK Then
            Set AccessAppFromHwnd = ApplicationFromComObject(obj)
            If Not AccessAppFromHwnd Is Nothing Then Exit Function
        End If
        hwndChild = FindWindowExA(hwnd, hwndChild, vbNullString, vbNullString)
    Loop

End Function


'---------------------------------------------------------------------------------------
' Procedure : ApplicationFromComObject
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : The native OM may be Application or a Window; always return Application.
'---------------------------------------------------------------------------------------
'
Private Function ApplicationFromComObject(obj As Object) As Object
    If obj Is Nothing Then Exit Function
    LogUnhandledErrors
    On Error Resume Next
    Set ApplicationFromComObject = obj.Application
    If ApplicationFromComObject Is Nothing Then Set ApplicationFromComObject = obj
    If Err Then
        Err.Clear
        Set ApplicationFromComObject = Nothing
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExtractOpenDatabaseFromCommandLine
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Return the last token in a process command line that names a database
'           : file. Tokens are split on spaces outside of double quotes, so an
'           : unquoted path (a path with no spaces in it) is found as well as a
'           : quoted one. An empty return does not mean no database is open: COM
'           : OpenCurrentDatabase never reaches the command line.
'---------------------------------------------------------------------------------------
'
Public Function ExtractOpenDatabaseFromCommandLine(strCmd As String) As String

    Dim lngPos As Long
    Dim strChar As String
    Dim strToken As String
    Dim blnInQuotes As Boolean

    ' One past the end flushes the final token without duplicating the test below.
    For lngPos = 1 To Len(strCmd) + 1
        If lngPos > Len(strCmd) Then strChar = " " Else strChar = Mid$(strCmd, lngPos, 1)
        If strChar = """" Then
            blnInQuotes = Not blnInQuotes
        ElseIf (strChar = " " Or strChar = vbTab) And Not blnInQuotes Then
            If Len(strToken) > 0 Then
                If IsDatabaseFileName(strToken) Then ExtractOpenDatabaseFromCommandLine = strToken
                strToken = vbNullString
            End If
        Else
            strToken = strToken & strChar
        End If
    Next lngPos

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsDatabaseFileName
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : True when the path carries an extension Access opens as a database.
'---------------------------------------------------------------------------------------
'
Private Function IsDatabaseFileName(strPath As String) As Boolean
    Select Case LCase$(FSO.GetExtensionName(strPath))
        Case "accdb", "accda", "accde", "accdr", "accdt", "mdb", "mda", "mde", "adp", "ade"
            IsDatabaseFileName = True
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsFolderTrusted
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : True when strFolder is an Access trusted location, or sits under one
'           : that allows subfolders. Checks HKCU/HKLM user and policy hives.
'           : CurrentProject.IsTrusted is not used: Enable Content does not survive
'           : a new Access process, which is what /cmd INSTALL launches.
'---------------------------------------------------------------------------------------
'
Public Function IsFolderTrusted(strFolder As String) As Boolean

    Dim strNorm As String
    Dim strVer As String

    If Len(strFolder) = 0 Then Exit Function
    strNorm = StripSlash(strFolder) & PathSep
    strVer = Application.Version

    If TrustedLocationHiveCoversFolder(&H80000001, _
        "Software\Microsoft\Office\" & strVer & "\Access\Security\Trusted Locations", strNorm) Then
        IsFolderTrusted = True
        Exit Function
    End If
    If TrustedLocationHiveCoversFolder(&H80000001, _
        "Software\Policies\Microsoft\Office\" & strVer & "\Access\Security\Trusted Locations", strNorm) Then
        IsFolderTrusted = True
        Exit Function
    End If
    If TrustedLocationHiveCoversFolder(&H80000002, _
        "Software\Policies\Microsoft\Office\" & strVer & "\Access\Security\Trusted Locations", strNorm) Then
        IsFolderTrusted = True
        Exit Function
    End If
    If TrustedLocationHiveCoversFolder(&H80000002, _
        "Software\Microsoft\Office\" & strVer & "\Access\Security\Trusted Locations", strNorm) Then
        IsFolderTrusted = True
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : TrustedLocationHiveCoversFolder
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Enumerate trusted-location subkeys under a registry hive/key and return
'           : True if any of them covers strFolderNorm (already slash-terminated).
'---------------------------------------------------------------------------------------
'
Private Function TrustedLocationHiveCoversFolder(lngHive As Long, strKey As String, _
    strFolderNorm As String) As Boolean

    Dim objReg As Object
    Dim varSubKeys As Variant
    Dim varKey As Variant
    Dim strPath As String
    Dim lngAllow As Long
    Dim strTrusted As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Set objReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
    If objReg Is Nothing Then Exit Function

    objReg.EnumKey lngHive, strKey, varSubKeys
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    If IsEmpty(varSubKeys) Then Exit Function

    For Each varKey In varSubKeys
        strPath = vbNullString
        lngAllow = 0
        objReg.GetStringValue lngHive, strKey & "\" & CStr(varKey), "Path", strPath
        objReg.GetDWORDValue lngHive, strKey & "\" & CStr(varKey), "AllowSubfolders", lngAllow
        If Err Then Err.Clear
        If Len(strPath) > 0 Then
            strTrusted = StripSlash(strPath) & PathSep
            If lngAllow <> 0 Then
                If InStr(1, strFolderNorm, strTrusted, vbTextCompare) = 1 Then
                    TrustedLocationHiveCoversFolder = True
                    Exit Function
                End If
            Else
                If StrComp(strFolderNorm, strTrusted, vbTextCompare) = 0 Then
                    TrustedLocationHiveCoversFolder = True
                    Exit Function
                End If
            End If
        End If
    Next varKey

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetRebuildStatusFilePath
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Path of the rebuild status JSON under the source folder's logs/ directory.
'---------------------------------------------------------------------------------------
'
Public Function GetRebuildStatusFilePath(strSourceFolder As String) As String
    GetRebuildStatusFilePath = StripSlash(strSourceFolder) & PathSep & "logs" & PathSep & "rebuild-status.json"
End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteRebuildStatusFile
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Write (or update) the flat rebuild-status.json the agent polls after
'           : Access exits. Preserves phaseStarted from an existing file when the
'           : caller does not supply one.
'---------------------------------------------------------------------------------------
'
Public Sub WriteRebuildStatusFile(strFile As String, strStatus As String, _
    Optional strError As String, Optional strBuildLog As String, _
    Optional strPhaseStarted As String)

    Dim dStatus As Dictionary
    Dim dExisting As Dictionary
    Dim strStarted As String

    If Len(strFile) = 0 Then Exit Sub

    strStarted = strPhaseStarted
    If Len(strStarted) = 0 Then
        Set dExisting = ReadRebuildStatusFile(strFile)
        If Not dExisting Is Nothing Then
            If dExisting.Exists("phaseStarted") Then strStarted = Nz(dExisting("phaseStarted"), vbNullString)
        End If
    End If
    If Len(strStarted) = 0 Then strStarted = Format$(Now, "yyyy-mm-dd hh:nn:ss")

    Set dStatus = New Dictionary
    dStatus.Add "status", strStatus
    dStatus.Add "error", strError
    dStatus.Add "buildLog", strBuildLog
    dStatus.Add "phaseStarted", strStarted
    dStatus.Add "updated", Format$(Now, "yyyy-mm-dd hh:nn:ss")

    VerifyPath strFile
    WriteFile ConvertToJson(dStatus, JSON_WHITESPACE), strFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReadRebuildStatusFile
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Parse rebuild-status.json into a Dictionary, or Nothing if missing/invalid.
'---------------------------------------------------------------------------------------
'
Public Function ReadRebuildStatusFile(strFile As String) As Dictionary

    Dim strJson As String

    If Len(strFile) = 0 Then Exit Function
    If Not FSO.FileExists(strFile) Then Exit Function

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    strJson = ReadFile(strFile)
    If Len(strJson) = 0 Then Exit Function
    Set ReadRebuildStatusFile = ParseJson(strJson)
    CatchAny eelError, "Unable to parse rebuild status file", _
        ModuleName & ".ReadRebuildStatusFile", True, True

End Function


'---------------------------------------------------------------------------------------
' Procedure : RecordSilentInstallResult
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : During /cmd INSTALL SILENT, write the terminal status using the source
'           : path saved by AfterBuild. No-ops when that path is missing.
'---------------------------------------------------------------------------------------
'
Private Sub RecordSilentInstallResult(strStatus As String, Optional strError As String)

    Dim strSource As String

    If Operation.InteractionMode <> eimSilent Then Exit Sub
    strSource = GetSetting(PROJECT_NAME, "Install", "Source Path", vbNullString)
    If Len(strSource) = 0 Then Exit Sub
    WriteRebuildStatusFile GetRebuildStatusFilePath(strSource), strStatus, strError

End Sub


'---------------------------------------------------------------------------------------
' Procedure : OpenAddinFile
' Author    : hecon5
' Date      : 1/15/2021
' Purpose   : runs a script to complete the addin trusting process. Once a trusted
'           : location is set, the file needs to be opened to trust it in many
'           : Corporate / Government environments due to security concerns.
'           : This will complete the process without the user needing to know
'           : where the file resides.
'           : It waits for two files to close (the "installer" and the "addin".
'           : This should hopefully ensure Access was closed prior to relaunch and
'           : significantly reduces instance of the application
'           : The subroutine is private because if you have called the addin from
'           : somewhere (aka, you're not installing it), opening the same file twice
'           : will cause headaches and likely corrupt the file.
'---------------------------------------------------------------------------------------
Public Sub OpenAddinFile(strAddInFileName As String, _
                            strInstallerFileName As String)

    Dim strScriptFile As String
    Dim strExt As String
    Dim lockFilePathAddin As String
    Dim lockFilePathInstaller As String

    ' Build file paths for lock files and batch script
    strExt = "." & FSO.GetExtensionName(strInstallerFileName)
    lockFilePathAddin = Replace(strAddInFileName, strExt, ".laccdb", , , vbTextCompare)
    lockFilePathInstaller = Replace(strInstallerFileName, strExt, ".laccdb", , , vbTextCompare)
    strScriptFile = Replace(strAddInFileName, strExt, ".cmd", , , vbTextCompare)

    ' Build batch script content
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "@Echo Off"
        .Add "setlocal ENABLEDELAYEDEXPANSION"
        .Add "ECHO Waiting for Addin file to copy over..."
        .Add ":WAITFORADDIN"
        .Add "ping 127.0.0.1 -n 1 -w 100 > nul"
        .Add "SET /a counter+=1"
        .Add "IF !counter!==300 GOTO DONE"
        .Add "IF NOT EXIST """, strAddInFileName, """ GOTO WAITFORADDIN"
        .Add "ECHO Waiting for Access to close..."
        .Add "SET /a counter=0"
        .Add ":WAITCLOSEINSTALLER"
        .Add "ping 127.0.0.1 -n 1 -w 100 > nul"
        .Add "SET /a counter+=1"
        .Add "IF !counter!==30 GOTO WAITCLOSEADDIN"
        .Add "IF EXIST """, lockFilePathInstaller, """ GOTO WAITCLOSEINSTALLER"
        .Add ":WAITCLOSEADDIN"
        .Add "ping 127.0.0.1 -n 1 -w 100 > nul"
        .Add "IF !counter!==40 GOTO MOVEON"
        .Add "IF EXIST """, lockFilePathAddin, """ GOTO WAITCLOSEADDIN"
        .Add ":OPENADDIN"
        .Add "ECHO Opening Add-in to finish installation..."
        .Add "ECHO (This window will automatically close when complete.)"
        .Add """", strAddInFileName, """"
        .Add "GOTO DONE"
        .Add ":MOVEON"
        .Add "Del """, lockFilePathAddin, """"
        .Add "Del """, lockFilePathInstaller, """"
        .Add "GOTO OPENADDIN"
        .Add ":DONE"
        .Add "Del """, strScriptFile, """"

        ' Write to file
        WriteFile .GetStr, strScriptFile
    End With

    ' Execute script
    Shell strScriptFile, vbNormalFocus

End Sub


'---------------------------------------------------------------------------------------
' Procedure : VerifyActivexNotDisabled
' Author    : Adam Waller
' Date      : 4/14/2023
' Purpose   : Verify that ActiveX has not been disabled in the registry, and warn the
'           : user that the add-in may not be able to build from source without this.
'---------------------------------------------------------------------------------------
'
Public Sub VerifyActivexNotDisabled()
    If IsActivexDisabled Then
        MsgBox2 "ActiveX Disabled", "WARNING: ActiveX appears to be disabled in the " & _
            "Microsoft Office Trust Center settings, or by a Group Policy setting. " & _
            "Microsoft Access uses ActiveX when importing content from XML, so some features " & _
            "of this add-in, such as building from source may not work " & _
            "correctly without enabling ActiveX.", _
            "You may need to review the ActiveX security settings  with your IT Department " & _
            "or system administrator to determine the appropriate setting for your system.", vbExclamation
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsActivexDisabled
' Author    : Adam Waller
' Date      : 4/14/2023
' Purpose   : Returns true if ActiveX appears to be enabled on the current system.
'           : (ActiveX is required to import XML files, such as table definitions when
'           :  building a database from source.) See issue #396
'---------------------------------------------------------------------------------------
'
Private Function IsActivexDisabled() As Boolean
    IsActivexDisabled = Not ( _
        CheckRegKey("HKCU\SOFTWARE\Policies\Microsoft\Office\common\security\disableallactivex", 0, Null) And _
        CheckRegKey("HKCU\SOFTWARE\Microsoft\Office\Common\Security\disableallactivex", 0, Null) And _
        CheckRegKey("HKCU\SOFTWARE\Policies\Microsoft\Office\" & Application.Version & "\Common\com categories\checkofficeactivex", 0, 1, Null) And _
        CheckRegKey("HKCU\SOFTWARE\Microsoft\Office\" & Application.Version & "\Common\com categories\checkofficeactivex", 0, 1, Null))
End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckRegKey
' Author    : Adam Waller
' Date      : 4/14/2023
' Purpose   : Check a registry key for specific allowed values, (including null)
'---------------------------------------------------------------------------------------
'
Private Function CheckRegKey(strPath As String, ParamArray AllowedValues() As Variant) As Boolean

    Dim varValue As Variant
    Dim intCnt As Integer

    LogUnhandledErrors
    On Error Resume Next

    ' Attempt to read registry key
    With New IWshRuntimeLibrary.WshShell
        varValue = .RegRead(strPath)
        ' A file not found error means the key did not exist.
        If Catch(-2147024894) Then varValue = Null
    End With

    ' Compare to array of allowed values
    For intCnt = 0 To UBound(AllowedValues)
        If varValue = AllowedValues(intCnt) Or _
            (IsNull(varValue) And IsNull(AllowedValues(intCnt))) Then
            CheckRegKey = True
            Exit For
        End If
    Next intCnt

    If Err Then Err.Clear

End Function


'---------------------------------------------------------------------------------------
' Procedure : IncrementAppVersion
' Author    : Adam Waller
' Date      : 1/6/2017
' Purpose   : Increments the build version (1.0.12)
'---------------------------------------------------------------------------------------
'
Public Sub IncrementAppVersion(Optional ReleaseType As eReleaseType = Build_xxV)

    Dim varParts As Variant
    Dim strFrom As String

    If ReleaseType = Same_Version Then Exit Sub
    strFrom = AppVersion
    varParts = Split(AppVersion, ".")
    varParts(ReleaseType) = varParts(ReleaseType) + 1
    If ReleaseType < Minor_xVx Then varParts(Minor_xVx) = 0
    If ReleaseType < Build_xxV Then varParts(Build_xxV) = 0
    AppVersion = Join(varParts, ".")

    ' Display old and new versions
    Debug.Print "Updated from " & strFrom & " to " & AppVersion

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AppVersion
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Get the version from the database property.
'---------------------------------------------------------------------------------------
'
Public Property Get AppVersion() As String
    Dim strVersion As String
    strVersion = GetDBProperty("AppVersion", CodeDb)
    If strVersion = vbNullString Then strVersion = "1.0.0"
    AppVersion = strVersion
End Property


'---------------------------------------------------------------------------------------
' Procedure : AppVersion
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Set version property in current database.
'---------------------------------------------------------------------------------------
'
Public Property Let AppVersion(strVersion As String)
    SetDBProperty "AppVersion", strVersion, , CodeDb
End Property


'---------------------------------------------------------------------------------------
' Procedure : InstalledVersion
' Author    : Adam Waller
' Date      : 4/21/2020
' Purpose   : Returns the installed version of the add-in from the registry.
'           : (We are saving this in the user hive, since it requires admin rights
'           :  to change the keys actually used by Access to register the add-in)
'---------------------------------------------------------------------------------------
'
Public Property Let InstalledVersion(strVersion As String)
    SaveSetting PROJECT_NAME, "Add-in", "Installed Version", strVersion
End Property
Public Property Get InstalledVersion() As String
    InstalledVersion = GetSetting(PROJECT_NAME, "Add-in", "Installed Version", vbNullString)
End Property
