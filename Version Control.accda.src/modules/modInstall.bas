﻿Attribute VB_Name = "modInstall"
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


' Registry hive
Private Enum eHive
    ehHKLM
    ehHKCU
End Enum

' Used to determine if Access is running as administrator. (Required for installing the add-in)
Private Declare PtrSafe Function IsUserAnAdmin Lib "shell32" () As Long

Private Const ModuleName As String = "modInstall"

' Used to add a trusted location for the add-in path (when necessary)
Private Const mcstrTrustedLocationName = PROJECT_NAME & " Version Control"

' Use a private type to manage install settings.
Public Type udtInstallSettings
    blnTrustAddInFolder As Boolean
    blnUseRibbonAddIn As Boolean
    blnOpenAfterInstall As Boolean
    strInstallFolder As String
    blnSettingsLoaded As Boolean
End Type
Private this As udtInstallSettings


'---------------------------------------------------------------------------------------
' Procedure : AutoRun
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : This code runs when the add-in file is opened directly. It provides the
'           : user an easy way to update the add-in on their system.
'---------------------------------------------------------------------------------------
'
Public Function AutoRun() As Boolean

    ' See if the we are opening the file from the installed location.
    If CodeProject.FullName = GetAddInFileName Then

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
' Procedure : InstallVCSAddin
' Author    : Adam Waller
' Date      : 10/19/2020
' Purpose   : Installs/updates the add-in for the current user.
'           : Returns true if successful.
'---------------------------------------------------------------------------------------
'
Public Sub InstallVCSAddin(blnTrustFolder As Boolean, blnUseRibbon As Boolean, blnOpenAfterInstall As Boolean, strInstallFolder As String, _
            Optional ByVal blnCreateCompiledVersion As Boolean = False)

    Const OPEN_MODE_OPTION As String = "Default Open Mode for Databases"

    Dim strSource As String
    Dim strDest As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Load install settings from registry, then update with parameter values
    GetInstallSettings
    With this
        .blnUseRibbonAddIn = blnUseRibbon
        .blnOpenAfterInstall = blnOpenAfterInstall
        .blnTrustAddInFolder = blnTrustFolder
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
        Exit Sub
    End If

    ' Check default database open mode.
    If Application.GetOption(OPEN_MODE_OPTION) = 1 Then
        If MsgBox2("Default Open Mode set to Exclusive", _
            "The default open mode option for Microsoft Access is currently set to open databases in Exclusive mode by default. " & vbCrLf & _
            "This add-in needs to be opened in shared mode in order to install successfully.", _
            "Change the default open mode to 'Shared'?", vbYesNo + vbExclamation) = vbYes Then
            Application.SetOption OPEN_MODE_OPTION, 0
            MsgBox2 "Default Option Changed", _
                "Please restart Microsoft Access and run the install again.", , vbInformation
        End If
        Exit Sub
    End If

    ' Run any applicable upgrades
    RunUpgrades

    ' Verify the trusted location
    If this.blnTrustAddInFolder Then VerifyTrustedLocation

    ' Copy the add-in file
    If Not UpdateAddInFile(blnCreateCompiledVersion) Then Exit Sub

    ' Install the ribbon
    If this.blnUseRibbonAddIn Then
        ' Ensure that the ribbon is installed
        modCOMAddIn.VerifyComAddIn
    Else
        ' Remove if currently installed
        modCOMAddIn.UninstallComAddIn
    End If

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
        "of Microsoft Access before using the add-in.", vbInformation, "Version Control Add-in"

    ' Open add-in from installed location if required.
    If this.blnOpenAfterInstall Then OpenAddinFile GetAddInFileName, CodeProject.FullName

    ' Close Access after installation is complete.
    DoCmd.Quit

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

    ' Ask the user if they want to preserve their user settings.
    intResponse = MsgBox2("Save User Settings", "Would you like your user settings/options preserved?", _
        "Click YES to save these items so they can be used if you reinstall the add-in," & vbCrLf & _
        "Or click NO to remove all settings related to this add-on.", vbQuestion + vbYesNoCancel)

    ' Allow user to cancel if they are not sure how to answer the above prompt.
    If intResponse = vbCancel Then Exit Sub

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
    If DebugMode(False) Then On Error GoTo 0 Else On Error Resume Next

    ' Remove trusted location added by this add-in. (if found)
    RemoveTrustedLocation

    ' Remove COM add-in
    modCOMAddIn.UninstallComAddIn

    ' Remove On Save hook
    'modExportOnSaveHook.Uninstall

    ' Notify the user of the completion of the uninstall process.
    MsgBox2 "Success!", "Version Control System has now been uninstalled.", _
        "Microsoft Access will be closed to remove the remaining files.", _
        vbInformation, "Version Control Add-in"

    ' Use the worker script to actually remove the add-in files.
    ' (They cannot be removed when they are in use, such as when procesing the uninstall.)
    Worker.Run_UninstallAddin

End Sub


'---------------------------------------------------------------------------------------
' Procedure : UpdateAddInFile
' Author    : Adam Waller
' Date      : 5/22/2023
' Purpose   : Update the add-in database file. Return true if successful.
'---------------------------------------------------------------------------------------
'
Private Function UpdateAddInFile(ByVal blnCreateCompiledVersion As Boolean) As Boolean

    ' Make sure the destination folder exists
    VerifyPath GetAddInFileName

    ' Update the file
    LogUnhandledErrors
    On Error Resume Next
    If FSO.FileExists(GetAddInFileName) Then DeleteFile GetAddInFileName, True

    If blnCreateCompiledVersion Then
        CreateAccde CodeProject.FullName, GetAddInFileName
    Else
        FSO.CopyFile CodeProject.FullName, GetAddInFileName, True
    End If

    If Err Then
        MsgBox2 "Unable to Update File", _
            "Encountered error " & Err.Number & ": " & Err.Description & " when copying file.", _
            "Is the Version Control Add-in loaded in another instance of Microsoft Access?" & vbCrLf & _
            "Please check to be sure that the following file is not in use:" & _
            vbCrLf & GetAddInFileName, vbExclamation
        Err.Clear
    Else
        ' Copied file with no errors.
        UpdateAddInFile = True
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : CreateAccde
' Author    : Josef Poetzl
' Date      : 2/21/2025
' Purpose   : Create a compiled Access file
'---------------------------------------------------------------------------------------
'
Private Sub CreateAccde(ByVal strSourceFilePath As String, ByVal strDestFilePath As String)

    Dim strFileToCompile As String

    strFileToCompile = strDestFilePath & ".accdb"
    FSO.CopyFile strSourceFilePath, strFileToCompile, True

    ' use new Access instance to create accde
    With New Access.Application
        .Visible = True
        .SysCmd 603, (strFileToCompile), (strDestFilePath)
    End With

    FSO.DeleteFile strFileToCompile, True

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
Public Function GetAddInFileName() As String
    GetAddInFileName = FSO.BuildPath(GetInstallSettings.strInstallFolder, CodeProject.Name)
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
        .RegWrite strPath & "Library", GetAddInFileName, "REG_SZ"
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
    ShellEx FSO.BuildPath(SysCmd(acSysCmdAccessDir), "msaccess.exe"), """" & GetAddInFileName & """", "runas"
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

    ' Update project description
    VBE.ActiveVBProject.Description = "Version " & AppVersion & " deployed on " & Date

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
        strOldPath = BuildPath2(Environ$("AppData"), "Microsoft", "AddIns", CodeProject.Name)

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
    strOldPath = FSO.BuildPath(CodeProject.Path, FSO.GetBaseName(CodeProject.Name)) & ".json"
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
            .blnOpenAfterInstall = GetSetting(PROJECT_NAME, "Install", "Open File", CInt(False))
            .strInstallFolder = GetSetting(PROJECT_NAME, "Install", "Install Folder", DefaultAddInFolderPath)
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
        SaveSetting PROJECT_NAME, "Install", "Open File", CInt(.blnOpenAfterInstall)
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
