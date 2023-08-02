Attribute VB_Name = "modCOMAddIn"
'---------------------------------------------------------------------------------------
' Module    : modCOMAddIn
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Functions to handling the installing, removing, and verifying the
'           : TwinBasic COM add-in used to provide the ribbon interface.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

Private Const ModuleName As String = "modComAddIn"

' Constants for registry entries
Private Const cstrAddinFriendlyName As String = "Ribbon integration for MSAccessVCS add-in"


'---------------------------------------------------------------------------------------
' Procedure : VerifyComAddin
' Author    : Adam Waller
' Date      : 3/1/2022
' Purpose   : Verify that the ribbon add-in is installed and the latest version.
'---------------------------------------------------------------------------------------
'
Public Sub VerifyComAddIn()

    Dim strPath As String
    Dim strFile As String
    Dim strKey As String
    Dim strHash As String
    Dim blnUpdateRibbon As Boolean
    Dim blnInstall As Boolean

    ' Build path to ribbon folder
    strPath = GetAddInPath

    ' Ribbon XML file
    strFile = strPath & "Ribbon.xml"
    strKey = "Ribbon XML"
    If Not FSO.FileExists(strFile) Then
        modResource.ExtractResource strKey, strPath
    Else
        ' In the future we may allow the user to choose whether
        ' to keep their existing ribbon.
        strHash = modResource.GetResourceHash(strKey)
        If strHash <> vbNullString Then
            If strHash <> GetFileHash(strFile) Then
                modResource.ExtractResource strKey, strPath
                blnUpdateRibbon = True
            End If
        End If
    End If

    ' COM Add-in
    strFile = strPath & GetComAddInFileName
    strKey = "COM Addin x" & GetOfficeBitness

    ' Verify add-in file
    If Not FSO.FileExists(strFile) Then
        blnInstall = True
    Else
        ' Compare to embedded resource file
        strHash = modResource.GetResourceHash(strKey)
        If strHash <> vbNullString Then
            ' Reinstall if the file is different
            If strHash <> GetFileHash(strFile) Then blnInstall = True
        End If
    End If

    ' Verify COM registration
    If Not blnInstall Then blnInstall = Not DllIsRegistered

    ' Install/reinstall if needed
    If blnInstall Then
        ' Unload the add-in, so we don't try to overwrite a file that is in use
        UnloadAddIn
        RemoveComDll
        ' Extract the new file from the resources table
        modResource.ExtractResource strKey, strPath
        ' Register the add-in file
        RegisterCOMAddIn
        ' Now we should be able to load the add-in
        LoadAddIn
    Else
        If blnUpdateRibbon Then
            ' Reload the add-in to refresh the ribbon
            UnloadAddIn
            LoadAddIn
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReloadRibbon
' Author    : Adam Waller
' Date      : 4/1/2022
' Purpose   : Reloads the Fluent UI ribbon interface to reflect changes to the XML
'           : file, such as switching to a different language.
'---------------------------------------------------------------------------------------
'
Public Sub ReloadRibbon()
    UnloadAddIn
    LoadAddIn
End Sub


'---------------------------------------------------------------------------------------
' Procedure : UninstallComAddIn
' Author    : Adam Waller
' Date      : 4/5/2022
' Purpose   : Unload, unregister, and remove COM add-in.
'---------------------------------------------------------------------------------------
'
Public Sub UninstallComAddIn()

    Dim strPath As String

    ' Unload the add-in ribbon
    UnloadAddIn

    ' Unregister the DLL from the registry
    DllUnregisterServer

    ' Remove DLL file
    RemoveComDll

    ' Remove ribbon XML file
    strPath = GetAddInPath & "Ribbon.xml"
    If FSO.FileExists(strPath) Then DeleteFile strPath

    ' Update the list of COM add-ins
    Application.COMAddIns.Update

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveComDll
' Author    : Adam Waller
' Date      : 4/8/2022
' Purpose   : This can be a little tricky because if the COM add-in was loaded in
'           : Microsoft Access, it may have a handle open that prevents us from deleting
'           : the file. If we can't delete it, we can rename it to a temp file in the
'           : current user's temp folder.
'---------------------------------------------------------------------------------------
'
Private Sub RemoveComDll()

    Dim strPath As String
    Dim strTemp As String

    ' Build expected path for DLL
    strPath = GetAddInPath & GetComAddInFileName
    If FSO.FileExists(strPath) Then

        ' Attempt to delete it first
        LogUnhandledErrors
        On Error Resume Next
        DeleteFile strPath
        If Catch(70) Then
            ' File handle in use. Rename to temp file
            strTemp = GetTempFile
            DeleteFile strTemp
            FSO.MoveFile strPath, strTemp
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAddInPath
' Author    : Adam Waller
' Date      : 3/11/2022
' Purpose   : Return path to add-in installation folder
'---------------------------------------------------------------------------------------
'
Private Function GetAddInPath() As String
    GetAddInPath = GetInstallSettings.strInstallFolder & PathSep
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetComAddInFileName
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Return the file name for the COM add-in
'---------------------------------------------------------------------------------------
'
Private Function GetComAddInFileName() As String
    GetComAddInFileName = Replace("MSAccessVCSLib_winXX.dll", "XX", GetOfficeBitness)
End Function


'---------------------------------------------------------------------------------------
' Procedure : RegisterCOMAddIn
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Register the add-in in the Windows registry.
'---------------------------------------------------------------------------------------
'
Private Function RegisterCOMAddIn() As Boolean

    ' Register with list of Access add-ins
    DllRegisterServer

    ' Refresh the list of add-ins from the registry
    Application.COMAddIns.Update

    ' Return true if we can find the loaded object
    RegisterCOMAddIn = Not (GetCOMAddIn Is Nothing)

End Function


'---------------------------------------------------------------------------------------
' Procedure : UnloadAddIn
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Unload the COM add-in, if found.
'---------------------------------------------------------------------------------------
'
Private Sub UnloadAddIn()
    Dim addVCS As COMAddIn
    Set addVCS = GetCOMAddIn
    If Not addVCS Is Nothing Then addVCS.Connect = False
    Application.COMAddIns.Update
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadAddIn
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Load (connect) the COM add-in
'---------------------------------------------------------------------------------------
'
Private Sub LoadAddIn()
    Dim addVCS As COMAddIn
    Set addVCS = GetCOMAddIn
    If addVCS Is Nothing Then
        ' Add-in not found. May need to be registered
    Else
        ' Load the add-in
        addVCS.Connect = True
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetCOMAddIn
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Return a reference to the VCS COM Add-In, if available
'---------------------------------------------------------------------------------------
'
Private Function GetCOMAddIn() As COMAddIn
    Dim addIn As COMAddIn
    For Each addIn In Application.COMAddIns
        If addIn.Description = cstrAddinFriendlyName Then
            Set GetCOMAddIn = addIn
            Exit For
        End If
    Next addIn
End Function


'---------------------------------------------------------------------------------------
' Procedure : DllIsRegistered
' Author    : Adam Waller
' Date      : 4/7/2022
' Purpose   : Checks for the CLSID registrations to verify that the COM DLL is properly
'           : registered in the registry for the current user.
'           : (This can change if the DLL is compiled in twinBASIC and registered to
'           :  the compiled DLL instead of the installed one.)
'---------------------------------------------------------------------------------------
'
Private Function DllIsRegistered() As Boolean

    Dim strTest As String

    ' Check HKLM registry key
    With New IWshRuntimeLibrary.WshShell
        ' We should have a value here if the install ran in the past.
        If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
        ' Look up the class ID from the COM registration
        strTest = .RegRead("HKCU\SOFTWARE\Classes\MSAccessVCSLib.AddInRibbon\CLSID\")
        If strTest <> vbNullString Then
            ' Read the file path for the registered DLL
            strTest = .RegRead("HKCU\SOFTWARE\Classes\CLSID\" & strTest & "\InProcServer32\")
            ' See if it matches the installation folder
            If strTest = GetAddInPath & GetComAddInFileName Then
                ' Path matches. See if the file actually exists
                DllIsRegistered = FSO.FileExists(strTest)
            End If
        End If
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : DllRegisterServer
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Register the add-in with the list of available add-ins for Access
'---------------------------------------------------------------------------------------
'
Private Sub DllRegisterServer()
    With New WshShell
        .Exec "regsvr32 /s """ & GetAddInPath & GetComAddInFileName & """"
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DllUnregisterServer
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Remove the add-in from the list
'---------------------------------------------------------------------------------------
'
Private Sub DllUnregisterServer()
    If Not DllIsRegistered Then Exit Sub
    With New WshShell
        .Exec "regsvr32 /u /s """ & GetAddInPath & GetComAddInFileName & """"
    End With
End Sub
