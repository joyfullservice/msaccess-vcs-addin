Option Compare Database
Option Explicit
Option Private Module

Public Enum eReleaseType
    Major_Vxx = 0
    Minor_xVx = 1
    Build_xxV = 2
    Same_Version = 3
End Enum

Private Enum eHive
    ehHKLM
    ehHKCU
End Enum

' Used to determine if Access is running as administrator. (Required for installing the add-in)
Private Declare PtrSafe Function IsUserAnAdmin Lib "shell32" () As Long

' Used to relaunch Access as an administrator to install the addin.
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As LongPtr

Private Const SW_SHOWNORMAL = 1


'---------------------------------------------------------------------------------------
' Procedure : AddInMenuItemLaunch
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Launch the main add-in form.
'---------------------------------------------------------------------------------------
'
Public Function AddInMenuItemLaunch() As Boolean
    PreloadVBE
    Form_frmVCSMain.Visible = True
    AddInMenuItemLaunch = True
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddInMenuItemExport
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Open main form and start export immediately. (Save users a click)
'---------------------------------------------------------------------------------------
'
Public Function AddInMenuItemExport() As Boolean
    PreloadVBE
    Form_frmVCSMain.Visible = True
    DoEvents
    Form_frmVCSMain.cmdExport_Click
    AddInMenuItemExport = True
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
    Dim strMsgBoxTitle As String
    Dim strMsgBoxText As String

    If CodeProject.FullName = GetAddinFileName Then
        ' Opening the file from add-in location, which would normally be unusual unless we are trying to remove
        ' legacy registry entries.
        If IsUserAnAdmin = 1 Then RemoveLegacyInstall
    Else
        ' Could be running it from another location, such as after downloading
        ' and updated version of the addin. In that case, we are either trying
        ' to install it for the first time, or trying to upgrade it.
        If IsAlreadyInstalled Then
            If InstalledVersion <> AppVersion Then
                strMsgBoxTitle = "Upgrade Version Control?"
                strMsgBoxText = "Would you like to upgrade to version " & AppVersion & "?"
            Else
                strMsgBoxTitle = "Reinstall Version Control?"
                strMsgBoxText = "Version " & AppVersion & " is already installed, would you like to reinstall it?"
            End If
            
            If MsgBox2(strMsgBoxTitle, strMsgBoxText, "Click 'Yes' to continue or 'No' to cancel.", vbQuestion + vbYesNo, "Version Control Add-in") = vbYes Then
                If InstallVCSAddin Then
                    MsgBox2 "Success!", "Version Control System add-in has been updated to " & AppVersion & ".", _
                        "Please restart any open instances of Microsoft Access before using the add-in.", vbInformation, "Version Control Add-in"
                    CheckForLegacyInstall
                    DoCmd.Quit
                End If
            Else
                ' Go to visual basic editor
                DoEvents
                DoCmd.RunCommand acCmdVisualBasicEditor
                DoEvents
            End If
        Else
            ' Not yet installed. Offer to install.
            If MsgBox2("Install Version Control?", _
                "Would you like to install version " & AppVersion & "?", _
                "Click 'Yes' to continue or 'No' to cancel.", vbQuestion + vbYesNo, "Version Control Add-in") = vbYes Then
                
                If InstallVCSAddin Then
                    MsgBox2 "Success!", "Version Control System has now been installed.", _
                        "You may begin using this tool after reopening Microsoft Access", vbInformation, "Version Control Add-in"
                    CheckForLegacyInstall
                End If
                
                DoCmd.Quit
            End If
        End If
    End If
    AutoRun = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : RunExportForCurrentDB
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : The primary purpose of this function is to be able to use VBA code to
'           : initiate a source code export, without currupting the current DB. This
'           : would typically be used in a build automation environment, or when
'           : exporting code from the add-in itself.
'           : To avoid causing file corruption issues, we need to run the export using
'           : the installed add-in, not the local MSAccessVCS project. In order to do
'           : this, we need to load the VCS add-in at the application level, then
'           : make it the active VB Project, then call the export function. When the
'           : export function is called, we need to complete any running code in the
'           : current database before export, so we will use a timer callback to
'           : launch the export cleanly from the installed add-in.
'           : This sounds complicated, but it is critical that we don't attempt to
'           : export code from a module that is currently running, or it may corrupt
'           : the file and cause Access to crash the next time the file is opened.
'           : (This can be repaired by rebuilding from source, but let's work to
'           :  prevent the problem in the first place.)
'---------------------------------------------------------------------------------------
'
Public Function RunExportForCurrentDB()

    ' Make sure the add-in is loaded.
    If Not AddinLoaded Then LoadVCSAddIn

    ' Set add-in project to active, just in case we are working
    ' on another development copy of the add-in.
    Set VBE.ActiveVBProject = GetAddInProject

    ' Call export function with an API callback.
    Run "LaunchExportAfterTimer"

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
    strAddInPath = Environ$("AppData") & "\Microsoft\AddIns\Version Control.accda"

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
        ' Set the active VB project so we can call the export function.
        Set VBE.ActiveVBProject = objAddIn
        
        ' Launch export for current database.
        ' (It is very important to use RunExportForCurrentDB when calling from the database from
        '  which you want to export source.)
        Application.Run "MSAccessVCS.RunExportForCurrentDB"
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddinLoaded
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : Returns true if the VCS add-in is currently loaded as a VBE Project.
'---------------------------------------------------------------------------------------
'
Private Function AddinLoaded() As Boolean
    AddinLoaded = Not GetAddInProject Is Nothing
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadVCSAddIn
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : Load the add-in at the application level so it can stay active
'           : even if the current database is closed.
'           : https://stackoverflow.com/questions/62270088/how-can-i-launch-an-access-add-in-not-com-add-in-from-vba-code
'---------------------------------------------------------------------------------------
'
Private Sub LoadVCSAddIn()
    ' The following lines will load the add-in at the application level,
    ' but will not actually call the function. Ignore the error of function not found.
    On Error Resume Next
    Application.Run GetAddinFileName & "!DummyFunction"
    On Error GoTo 0
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAddInProject
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : Return the VBProject of the MSAccessVCS add-in.
'---------------------------------------------------------------------------------------
'
Private Function GetAddInProject() As VBProject
    Dim oProj As VBProject
    For Each oProj In VBE.VBProjects
        If StrComp(oProj.FileName, GetAddinFileName, vbTextCompare) = 0 Then
            Set GetAddInProject = oProj
            Exit For
        End If
    Next oProj
End Function


'---------------------------------------------------------------------------------------
' Procedure : InstallVCSAddin
' Author    : Adam Waller
' Date      : 10/19/2020
' Purpose   : Installs/updates the add-in for the current user.
'           : Returns true if successful.
'---------------------------------------------------------------------------------------
'
Private Function InstallVCSAddin() As Boolean
    
    Const OPEN_MODE_OPTION As String = "Default Open Mode for Databases"
    
    Dim strSource As String
    Dim strDest As String
    
    strSource = CodeProject.FullName
    strDest = GetAddinFileName
    
    ' We can't replace a file with itself.  :-)
    If strSource = strDest Then Exit Function
    
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
        Exit Function
    End If
    
    ' Copy the file, overwriting any existing file.
    ' Requires FSO to copy open database files. (VBA.FileCopy give a permission denied error.)
    ' We also use FSO to force the deletion of the existing file, if found.
    On Error Resume Next
    If FSO.FileExists(strDest) Then DeleteFile strDest, True
    FSO.CopyFile strSource, strDest, True
    If Err Then
        MsgBox2 "Unable to update file", _
            "Encountered error " & Err.Number & ": " & Err.Description & " when copying file.", _
            "Please check to be sure that the following file is not in use:" & vbCrLf & strDest, vbExclamation
        Err.Clear
        On Error GoTo 0
    Else
        On Error GoTo 0
        ' Register the Menu controls
        RegisterMenuItem "&Version Control", "=AddInMenuItemLaunch()"
        RegisterMenuItem "&Export All Source", "=AddInMenuItemExport()"
        ' Update installed version number
        InstalledVersion = AppVersion
        ' Return success
        InstallVCSAddin = True
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : UninstallVCSAddin
' Author    : Adam Kauffman
' Date      : 5/27/2020
' Purpose   : Removes the add-in for the current user.
'           : Returns true if successful.
'---------------------------------------------------------------------------------------
'
Public Function UninstallVCSAddin() As Boolean
    
    Dim strDest As String
    strDest = GetAddinFileName
    
    ' Copy the file, overwriting any existing file.
    ' Requires FSO to copy open database files. (VBA.FileCopy give a permission denied error.)
    On Error Resume Next
    DeleteFile strDest, True
    On Error GoTo 0
    
    ' Error 53 = File Not found is okay.
    If Err.Number <> 0 And Err.Number <> 53 Then
        MsgBox2 "Unable to delete file", _
            "Encountered error " & Err.Number & ": " & Err.Description & " when copying file.", _
            "Please check to be sure that the following file is not in use:" & vbCrLf & strDest, vbExclamation
        Err.Clear
    Else
        ' Register the Menu controls
        RemoveMenuItem "&Version Control", "=AddInMenuItemLaunch()"
        RemoveMenuItem "&Export All Source", "=AddInMenuItemExport()"
        ' Update installed version number
        InstalledVersion = 0
        ' Return success
        UninstallVCSAddin = True
    End If
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetAddinFileName
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : This is where the add-in would be installed.
'---------------------------------------------------------------------------------------
'
Private Function GetAddinFileName() As String
    GetAddinFileName = Environ$("AppData") & "\Microsoft\AddIns\" & CodeProject.Name
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsAlreadyInstalled
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Returns true if the addin is already installed.
'---------------------------------------------------------------------------------------
'
Public Function IsAlreadyInstalled() As Boolean
    
    Dim strPath As String
    Dim strTest As String
    
    ' Check for registry key of installed version
    If InstalledVersion <> vbNullString Then
        
        ' Check for addin file
        If LCase(FSO.GetFileName(GetAddinFileName)) = LCase(CodeProject.Name) Then
            strPath = GetAddinRegPath & "&Version Control\Library"
            
            ' Check HKLM registry key
            With New IWshRuntimeLibrary.WshShell
                ' We should have a value here if the install ran in the past.
                On Error Resume Next
                strTest = .RegRead(strPath)
            End With
            
            If Err.Number > 0 Then Err.Clear
            On Error GoTo 0
            
            ' Return our determination
            IsAlreadyInstalled = (strTest <> vbNullString)
        End If
    End If
    
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
' Purpose   : Add the menu item through the registry (HKLM, requires admin)
'---------------------------------------------------------------------------------------
'
Private Sub RegisterMenuItem(ByVal strName As String, Optional ByVal strFunction As String = "=LaunchMe()")

    Dim strPath As String
    
    ' We need to create/update three registry keys for each item.
    strPath = GetAddinRegPath & strName & "\"
    With New IWshRuntimeLibrary.WshShell
        .RegWrite strPath & "Expression", strFunction, "REG_SZ"
        .RegWrite strPath & "Library", GetAddinFileName, "REG_SZ"
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
Private Sub RemoveMenuItem(ByVal strName As String, Optional ByVal strFunction As String = "=LaunchMe()", Optional Hive As eHive = ehHKCU)

    Dim strPath As String
    Dim objShell As WshShell
    
    ' We need to remove three registry keys for each item.
    strPath = GetAddinRegPath(Hive) & strName & "\"
    Set objShell = New WshShell
    With objShell
        ' Just in case someone changed some of the keys...
        On Error Resume Next
        .RegDelete strPath & "Expression"
        .RegDelete strPath & "Library"
        .RegDelete strPath & "Version"
        .RegDelete strPath
        If Err Then Err.Clear
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
    ShellExecute 0, "runas", SysCmd(acSysCmdAccessDir) & "\msaccess.exe", """" & GetAddinFileName & """", vbNullString, SW_SHOWNORMAL
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
    Dim blnSuccess As Boolean
    
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
    strBinaryFile = CodeProject.Path & "\Version_Control_v" & AppVersion & ".zip"
    If FSO.FileExists(strBinaryFile) Then DeleteFile strBinaryFile, True
    CreateZipFile strBinaryFile
    CopyFileToZip CodeProject.FullName, strBinaryFile
    
    ' Deploy latest version on this machine
    blnSuccess = InstallVCSAddin
    
    ' Use the newly installed add-in to Export the project to version control.
    RunExportForCurrentDB
    
    ' Finish with success message if the latest version was installed.
    If blnSuccess Then Debug.Print "Version " & AppVersion & " installed."
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IncrementAppVersion
' Author    : Adam Waller
' Date      : 1/6/2017
' Purpose   : Increments the build version (1.0.12)
'---------------------------------------------------------------------------------------
'
Public Sub IncrementAppVersion(ReleaseType As eReleaseType)
    Dim varParts As Variant
    If ReleaseType = Same_Version Then Exit Sub
    varParts = Split(AppVersion, ".")
    varParts(ReleaseType) = varParts(ReleaseType) + 1
    If ReleaseType < Minor_xVx Then varParts(Minor_xVx) = 0
    If ReleaseType < Build_xxV Then varParts(Build_xxV) = 0
    AppVersion = Join(varParts, ".")
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
    strVersion = GetDBProperty("AppVersion")
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
    SetDBProperty "AppVersion", strVersion
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
    SaveSetting GetCodeVBProject.Name, "Add-in", "Installed Version", strVersion
End Property
Public Property Get InstalledVersion() As String
    InstalledVersion = GetSetting(GetCodeVBProject.Name, "Add-in", "Installed Version", vbNullString)
End Property


'---------------------------------------------------------------------------------------
' Procedure : HasLegacyInstall
' Author    : Adam Waller
' Date      : 5/27/2020
' Purpose   : Returns true if legacy registry entries are found.
'---------------------------------------------------------------------------------------
'
Public Sub CheckForLegacyInstall()
    
    Dim strPath As String
    Dim strTest As String
    Dim objShell As IWshRuntimeLibrary.WshShell
    
    If InstalledVersion < "3.2.0" Then
        strPath = GetAddinRegPath(ehHKLM) & "&Version Control\Library"
        Set objShell = New IWshRuntimeLibrary.WshShell
        On Error Resume Next
        strTest = objShell.RegRead(strPath)
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
    RemoveMenuItem "&Version Control", "=AddInMenuItemLaunch()", ehHKLM
    RemoveMenuItem "&Export All Source", "=AddInMenuItemExport()", ehHKLM

    MsgBox2 "Legacy Items Removed", "Thanks for getting those cleaned up!" _
        , "Microsoft Access will now close so you can continue.", vbInformation
    DoCmd.Quit
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PreloadVBE
' Author    : Adam Waller
' Date      : 5/25/2020
' Purpose   : Force Access to load the VBE project. (This can help prevent crashes
'           : when code is run before the VB Project is fully loaded.)
'---------------------------------------------------------------------------------------
'
Public Sub PreloadVBE()
    Dim strName As String
    DoCmd.Hourglass True
    strName = VBE.ActiveVBProject.Name
    DoCmd.Hourglass False
End Sub