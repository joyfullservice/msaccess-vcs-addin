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
Private Const cstrAddinProjectName As String = "MSAccessVCS"
Private Const cstrAddinClassName As String = "AddInRibbon"
Private Const cstrAddinQualifiedClassName As String = cstrAddinProjectName & "." & cstrAddinClassName
Private Const cstrRootRegistryFolder As String = "HKCU\SOFTWARE\Microsoft\Office\Access\Addins\" & cstrAddinQualifiedClassName & "\"


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
    strPath = Environ$("AppData") & PathSep & "MSAccessVCS" & PathSep
    
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
        RegisterCOMAddIn strFile
        ' Now we should be able to load the add-in
        LoadAddIn strFile
    Else
        If blnUpdateRibbon Then
            ' Reload the add-in to refresh the ribbon
            UnloadAddIn
            LoadAddIn strFile
        End If
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAddInFileName
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Return the file name for the COM add-in
'---------------------------------------------------------------------------------------
'
Private Function GetAddInFileName() As String
    GetAddInFileName = Replace("MSAccessVCS_winXX.dll", "XX", GetOfficeBitness)
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
' Purpose   : Register the add-in in the Windows registry. We will add it to this
'           : project as a temporary reference to trigger the registration.
'---------------------------------------------------------------------------------------
'
Private Function RegisterCOMAddIn(strFile As String) As Boolean

    Dim proj As VBProject
    Dim ref As VBIDE.Reference
    Dim intCnt As Integer
    
    ' Use code project just in case this is run from the loaded Access add-in
    Set proj = GetCodeVBProject
    
    ' Add a temporary reference to the file, then remove it
    With proj.References
        intCnt = .Count
        Set ref = .AddFromFile(strFile)
        If .Count > intCnt Then
            .Remove ref
        Else
            Log.Error eelError, "Ribbon add-in registration failed for " & strFile, ModuleName & ".RegisterCOMAddIn"
        End If
    End With

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
Private Sub LoadAddIn(strPath As String)
    
    Dim addVCS As COMAddIn
    
    Set addVCS = GetCOMAddIn
    If addVCS Is Nothing Then
        'application.COMAddIns.
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
        If addIn.Description = "asdfsajdf" Then
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

    On Error GoTo RegError
    
    With New IWshRuntimeLibrary.WshShell
        .RegWrite cstrRootRegistryFolder & "FriendlyName", cstrAddinProjectName, "REG_SZ"
        .RegWrite cstrRootRegistryFolder & "Description", cstrAddinProjectName, "REG_SZ"
        .RegWrite cstrRootRegistryFolder & "LoadBehavior", 3, "REG_DWORD"
    End With
    
    DllRegisterServer = True
    Exit Function
    
RegError:
    MsgBox "DllRegisterServer -- An error occured trying to write to the system registry:" & vbCrLf & _
            Err.Description & " (" & Hex(Err.Number) & ")"

End Function


'---------------------------------------------------------------------------------------
' Procedure : DllUnregisterServer
' Author    : Adam Waller
' Date      : 3/5/2022
' Purpose   : Remove the add-in from the list
'---------------------------------------------------------------------------------------
'
Private Function DllUnregisterServer() As Boolean

    On Error GoTo RegError
    
    With New IWshRuntimeLibrary.WshShell
        .RegDelete cstrRootRegistryFolder & "FriendlyName"
        .RegDelete cstrRootRegistryFolder & "Description"
        .RegDelete cstrRootRegistryFolder & "LoadBehavior"
        .RegDelete cstrRootRegistryFolder
    End With
    
    DllUnregisterServer = True
    Exit Function
        
RegError:
        MsgBox "DllUnregisterServer -- An error occured trying to delete from the system registry:" & vbCrLf & _
                Err.Description & " (" & Hex(Err.Number) & ")"
        
End Function
