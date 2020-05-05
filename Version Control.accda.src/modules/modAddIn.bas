Option Compare Database
Option Explicit
Option Private Module

Public Enum eReleaseType
    Major_Vxx = 0
    Minor_xVx = 1
    Build_xxV = 2
End Enum

' Used to determine if Access is running as administrator. (Required for installing the add-in)
Private Declare PtrSafe Function IsUserAnAdmin Lib "shell32" () As Long

' Used to relaunch Access as an administrator to install the addin.
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Private Const SW_SHOWNORMAL = 1


'---------------------------------------------------------------------------------------
' Procedure : AddInMenuItemLaunch
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Launch the main add-in form.
'---------------------------------------------------------------------------------------
'
Public Function AddInMenuItemLaunch()
    Form_frmVCSMain.Visible = True
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddInMenuItemExport
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Open main form and start export immediately. (Save users a click)
'---------------------------------------------------------------------------------------
'
Public Function AddInMenuItemExport()
    Form_frmVCSMain.Visible = True
    DoEvents
    Form_frmVCSMain.cmdExport_Click
End Function


'---------------------------------------------------------------------------------------
' Procedure : AutoRun
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : This code runs when the add-in file is opened directly. It provides the
'           : user an easy way to update the add-in on their system.
'---------------------------------------------------------------------------------------
'
Public Function AutoRun()

    ' If we are running from the addin location, we might be trying to register it.
    If CodeProject.FullName = GetAddinFileName Then
    
        ' See if the user has admin privileges
        If IsUserAnAdmin = 1 Then
        
            ' Create the menu items
            ' NOTE: Be sure to keep these consistent with the USysRegInfo table
            ' so the user can uninstall the add-in later if desired.
            RegisterMenuItem "&Version Control", "=AddInMenuItemLaunch()"
            RegisterMenuItem "&Export All Source", "=AddInMenuItemExport()"
            InstalledVersion = AppVersion
            
            ' Give success message and quit Access
            If IsAlreadyInstalled Then
                MsgBox2 "Success!", "Version Control System has now been installed.", _
                    "You may begin using this tool after reopening Microsoft Access", vbInformation, "Version Control Add-in"
                DoCmd.Quit
            End If
        Else
            ' User does not have admin priviledges. Shouldn't normally be opening the add-in directly.
            ' Don't do anything special here. Just let them browse around in the file.
        End If
    Else
        ' Could be running it from another location, such as after downloading
        ' and updated version of the addin. In that case, we are either trying
        ' to install it for the first time, or trying to upgrade it.
        If IsAlreadyInstalled Then
            If InstalledVersion <> AppVersion Then
                If MsgBox2("Upgrade Version Control?", _
                    "Would you like to upgrade to version " & AppVersion & "?", _
                    "Click 'Yes' to continue or 'No' to cancel.", vbQuestion + vbYesNo, "Version Control Add-in") = vbYes Then
                    If InstallVCSAddin Then
                        MsgBox2 "Success!", "Version Control System add-in has been updated to " & AppVersion & ".", _
                            "Please restart any open instances of Microsoft Access before using the add-in.", vbInformation, "Version Control Add-in"
                        DoCmd.Quit
                    End If
                End If
            Else
                ' Go to visual basic editor, since that is the most likely destination.
                DoCmd.RunCommand acCmdVisualBasicEditor
            End If
        Else
            ' Not yet installed. Offer to install.
            If MsgBox2("Install Version Control?", _
                "Would you like to install version " & AppVersion & "?", _
                "Click 'Yes' to continue or 'No' to cancel.", vbQuestion + vbYesNo, "Version Control Add-in") = vbYes Then
                If InstallVCSAddin Then RelaunchAsAdmin
                DoCmd.Quit
            End If
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : InstallVCSAddin
' Author    : Adam Waller
' Date      : 1/14/2020
' Purpose   : Installs/updates the add-in for the current user.
'           : Returns true if successful.
'---------------------------------------------------------------------------------------
'
Private Function InstallVCSAddin()
    
    Dim strSource As String
    Dim strDest As String
    
    strSource = CodeProject.FullName
    strDest = GetAddinFileName
    
    ' We can't replace a file with itself.  :-)
    If strSource = strDest Then Exit Function
    
    ' Copy the file, overwriting any existing file.
    ' Requires FSO to copy open database files. (VBA.FileCopy give a permission denied error.)
    On Error Resume Next
    FSO.CopyFile strSource, strDest, True
    If Err Then
        MsgBox2 "Unable to update file", _
            "Encountered error " & Err.Number & ": " & Err.Description & " when copying file.", _
            "Please check to be sure that the following file is not in use:" & vbCrLf & strDest, vbExclamation
        Err.Clear
    Else
        ' Update installed version number
        InstalledVersion = AppVersion
        ' Return success
        InstallVCSAddin = True
    End If
    On Error GoTo 0

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAddinFileName
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : This is where the add-in would be installed.
'---------------------------------------------------------------------------------------
'
Private Function GetAddinFileName() As String
    GetAddinFileName = Environ("AppData") & "\Microsoft\AddIns\" & CodeProject.Name
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsAlreadyInstalled
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Returns true if the addin is already installed.
'---------------------------------------------------------------------------------------
'
Private Function IsAlreadyInstalled() As Boolean
    
    Dim strPath As String
    Dim oShell As IWshRuntimeLibrary.WshShell
    Dim strTest As String
    
    ' Check for registry key of installed version
    If InstalledVersion <> vbNullString Then
        
        ' Check for addin file
        If Dir(GetAddinFileName) = CodeProject.Name Then
            
            ' Check HKLM registry key
            Set oShell = New IWshRuntimeLibrary.WshShell
            strPath = GetAddinRegPath & "&Version Control\Library"
            On Error Resume Next
            ' We should have a value here if the install ran in the past.
            strTest = oShell.RegRead(strPath)
            If Err Then Err.Clear
            On Error GoTo 0
            Set oShell = Nothing
        
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
Private Function GetAddinRegPath() As String
    GetAddinRegPath = "HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\" & _
            Application.Version & "\Access\Menu Add-Ins\"
End Function


'---------------------------------------------------------------------------------------
' Procedure : RegisterMenuItem
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Add the menu item through the registry (HKLM, requires admin)
'---------------------------------------------------------------------------------------
'
Private Function RegisterMenuItem(strName, Optional strFunction As String = "=LaunchMe()")

    Dim oShell As IWshRuntimeLibrary.WshShell
    Dim strPath As String
    
    Set oShell = New IWshRuntimeLibrary.WshShell
    
    ' We need to create/update three registry keys for each item.
    strPath = GetAddinRegPath & strName & "\"
    With oShell
        .RegWrite strPath & "Expression", strFunction, "REG_SZ"
        .RegWrite strPath & "Library", GetAddinFileName, "REG_SZ"
        .RegWrite strPath & "Version", 3, "REG_DWORD"
    End With
    Set oShell = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : RelaunchAsAdmin
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Launch the addin file with admin privileges so the user can register it.
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
Public Sub Deploy(Optional ReleaseType As eReleaseType = Build_xxV)
    
    Const cstrSpacer As String = "--------------------------------------------------------------"
        
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
    
    ' Save all code modules
    DoCmd.RunCommand acCmdCompileAndSaveAllModules
    
    ' Export the source code to version control
    ExportSource
    
    ' Deploy latest version on this machine
    If InstallVCSAddin Then Debug.Print "Version " & AppVersion & " installed."
    
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
    If strVersion = "" Then strVersion = "1.0.0"
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
Private Property Let InstalledVersion(strVersion As String)
    SaveSetting GetCodeVBProject.Name, "Add-in", "Installed Version", strVersion
End Property
Private Property Get InstalledVersion() As String
    InstalledVersion = GetSetting(GetCodeVBProject.Name, "Add-in", "Installed Version", vbNullString)
End Property