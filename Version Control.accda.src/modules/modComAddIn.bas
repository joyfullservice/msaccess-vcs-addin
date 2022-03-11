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
    strFile = strPath & GetAddInFileName
    strKey = "COM Addin x" & GetOfficeBitness
    
    ' Verify add-in file
    If Not FSO.FileExists(strFile) Then
        blnInstall = True
    Else
        ' Compare to embedded resource file
        strHash = modResource.GetResourceHash(strKey)
        If strHash <> vbNullString Then
            ' Reinstall if the file is different
            blnInstall = (strHash <> GetFileHash(strFile))
        End If
    End If
    
    ' Install/reinstall if needed
    If blnInstall Then
        ' Unload the add-in, so we don't try to overwrite a file that is in use
        UnloadAddIn
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
' Procedure : GetAddInPath
' Author    : Adam Waller
' Date      : 3/11/2022
' Purpose   : Return path to add-in installation folder
'---------------------------------------------------------------------------------------
'
Private Function GetAddInPath() As String
    GetAddInPath = Environ$("AppData") & PathSep & "MSAccessVCS" & PathSep
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAddInFileName
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Return the file name for the COM add-in
'---------------------------------------------------------------------------------------
'
Private Function GetAddInFileName() As String
    GetAddInFileName = Replace("MSAccessVCSLib_winXX.dll", "XX", GetOfficeBitness)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetOfficeBitness
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Returns "32" or "64" as the bitness of Microsoft Office (not Windows)
'---------------------------------------------------------------------------------------
'
Private Function GetOfficeBitness() As String
    ' COM Add-in
    #If Win64 Then
        ' 64-bit add-in (Office x64)
        GetOfficeBitness = "64"
    #Else
        ' 32-bit add-in
        GetOfficeBitness = "32"
    #End If
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
' Procedure : DllRegisterServer
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Register the add-in with the list of available add-ins for Access
'---------------------------------------------------------------------------------------
'
Private Function DllRegisterServer() As Boolean
    With New WshShell
        .Exec "regsvr32 /s """ & GetAddInPath & GetAddInFileName & """"
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : DllUnregisterServer
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Remove the add-in from the list
'---------------------------------------------------------------------------------------
'
Private Function DllUnregisterServer() As Boolean
    With New WshShell
        .Exec "regsvr32 /u /s """ & GetAddInPath & GetAddInFileName & """"
    End With
End Function
