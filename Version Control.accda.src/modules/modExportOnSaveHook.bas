Attribute VB_Name = "modExportOnSaveHook"
'---------------------------------------------------------------------------------------
' Module    : modExportOnSaveHook
' Author    : bclothier
' Date      : 3/28/2023
' Purpose   : Handles hooking into and receiving callbacks from the hook DLL.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

' Must match the AccessObjectState definition used in hook's ObjectTracker module
Private Enum AccessObjectState
    ObjectIsNotOpenOrDoesntExist = 0
    ObjectIsOpen = Access.acObjStateOpen
    ObjectIsNew = Access.acObjStateNew
    ObjectIsDirty = Access.acObjStateDirty
    ObjectIsNewOrDirty = (Not ObjectIsOpen) ' Use in an AND operation and check if result is nonzero
End Enum

' Must match the HookConfiguration definition used in hook's DllGlobal module
Private Type HookConfiguration
    Size As Long
    App As Access.Application
    CallbackProject As VBIDE.VBProject
    AfterSaveRequestDelayMilliseconds As Long
    AfterSaveCallbackProcedureName As LongPtr
    LogFilePath As LongPtr
End Type

' Must match the ObjectData definition used in hook's ObjectTracker module
Private Type ObjectData
    Index As Integer
    Cancelled As Boolean
    ObjectType As Access.AcObjectType
    NewObjectType As Access.AcObjectType
    InitialObjectState As AccessObjectState
    DesignerhWnd As LongPtr
    OriginalName As String * 64
    NewName As String * 64
End Type

' HMODULE LoadLibraryExW(
'   [in] LPCWSTR lpLibFileName,
'        HANDLE  hFile,
'   [in] DWORD   dwFlags
' );
Private Declare PtrSafe Function LoadLibraryExW _
Lib "Kernel32.dll" ( _
    ByVal lpLibFileName As LongPtr, _
    ByVal hFile As LongPtr, _
    ByVal dwFlags As Long _
) As LongPtr

' BOOL FreeLibrary(
'   [in] HMODULE hLibModule
' );
Private Declare PtrSafe Function FreeLibrary _
Lib "Kernel32.dll" ( _
    ByVal hLibModule As LongPtr _
) As Long

#If Win64 Then
Private Declare PtrSafe Function StartHook _
Lib "MSAccessVCSHook_win64.dll" ( _
    ByRef Config As HookConfiguration _
) As Boolean

Private Declare PtrSafe Function StopHook _
Lib "MSAccessVCSHook_win64.dll" () As Boolean

#Else

Private Declare PtrSafe Function StartHook _
Lib "MSAccessVCSHook_win32.dll" ( _
    ByRef Config As HookConfiguration _
) As Boolean

Private Declare PtrSafe Function StopHook _
Lib "MSAccessVCSHook_win32.dll" () As Boolean

#End If

Private ptrLibraryHandle As LongPtr


'---------------------------------------------------------------------------------------
' Procedure : GetHookFileName
' Author    : bclothier
' Date      : 3/28/2023
' Purpose   : This is path where the hook DLL would be installed.
'---------------------------------------------------------------------------------------
'
Public Function GetHookFileName() As String
    GetHookFileName = Replace("MSAccessVCSHook_winXX.dll", "XX", GetOfficeBitness)
End Function


'---------------------------------------------------------------------------------------
' Procedure : VerifyHook
' Author    : bclothier
' Date      : 3/28/2023
' Purpose   : Verify that the hook is installed and the latest version.
'---------------------------------------------------------------------------------------
'
Public Sub VerifyHook()

    Dim strPath As String
    Dim strFile As String
    Dim strKey As String
    Dim strHash As String
    Dim blnInstall As Boolean

    ' Hook
    strPath = GetInstallSettings.strInstallFolder & PathSep
    strFile = FSO.BuildPath(strPath, GetHookFileName)
    strKey = "Hook x" & GetOfficeBitness()

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

    ' Install/reinstall if needed
    If blnInstall Then
        ' Extract the new file from the resources table
        modResource.ExtractResource strKey, strPath
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ActivateHook
' Author    : bclothier
' Date      : 3/28/2023
' Purpose   : Enables the hook to support automatic export after save
'---------------------------------------------------------------------------------------
'
Public Function ActivateHook(Optional ExportRequestDelayMilliseconds As Long = 500) As Boolean
    Dim Config As HookConfiguration

    If ptrLibraryHandle = 0 Then
        VerifyHook
        ptrLibraryHandle = LoadLibraryExW(StrPtr(FSO.BuildPath(GetInstallSettings.strInstallFolder, GetHookFileName)), &H0, &H0)
        If ptrLibraryHandle Then
            Config.Size = LenB(Config)
            Set Config.App = Application
            Set Config.CallbackProject = GetAddInProject
            If Config.CallbackProject Is Nothing Then
                With Application.WizHook
                    .Key = 51488399
                    Set Config.CallbackProject = .DbcVbProject
                End With
            End If
            Config.AfterSaveCallbackProcedureName = StrPtr("HandleExportCallback")
            Config.AfterSaveRequestDelayMilliseconds = ExportRequestDelayMilliseconds
            Config.LogFilePath = StrPtr(CodeProject.Path & "\MSAccessVCSHook.log")
            If StartHook(Config) = False Then
                FreeLibrary (ptrLibraryHandle)
                ptrLibraryHandle = 0
            End If
        End If
    End If

    ActivateHook = (ptrLibraryHandle <> 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : DeactivateHook
' Author    : bclothier
' Date      : 3/28/2023
' Purpose   : Deactivates the hook and unloads the library
'---------------------------------------------------------------------------------------
'
Public Function DeactivateHook() As Boolean
    If ptrLibraryHandle Then
        If StopHook() Then
            If FreeLibrary(ptrLibraryHandle) Then
                ptrLibraryHandle = 0
            End If
        End If
    End If

    DeactivateHook = (ptrLibraryHandle = 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : HandleExportCallback
' Author    : bclothier
' Date      : 3/28/2023
' Purpose   : Callback function from the hook DLL to provide information about modified
'           : objects that were saved.
'---------------------------------------------------------------------------------------
'
Public Sub HandleExportCallback(UpperBound As Long, ObjectDataArray() As ObjectData)
    On Error GoTo ErrHandler

    Dim dAccessObjects As Dictionary
    Dim oAccessObject As Access.AccessObject

    Dim strName As String

    Dim lngIndex As Long

    Set dAccessObjects = New Dictionary

    For lngIndex = 0 To UpperBound
        With ObjectDataArray(lngIndex)
            If .Cancelled = False Then
                If Len(Trim$(Replace$(.NewName, vbNullChar, vbNullString))) Then
                    strName = Trim$(.NewName)
                Else
                    strName = Trim$(.OriginalName)
                End If
                Select Case .NewObjectType
                    Case acTable
                        Set oAccessObject = CurrentData.AllTables(strName)
                    Case acTableDataMacro
                        Set oAccessObject = CurrentData.AllTables(strName)
                    Case acQuery
                        Set oAccessObject = CurrentData.AllQueries(strName)
                    Case acForm
                        Set oAccessObject = CurrentProject.AllForms(strName)
                    Case acReport
                        Set oAccessObject = CurrentProject.AllReports(strName)
                    Case acMacro
                        Set oAccessObject = CurrentProject.AllMacros(strName)
                    Case acModule
                        Set oAccessObject = CurrentProject.AllModules(strName)
                End Select

                dAccessObjects.Add .NewObjectType & "|" & strName, oAccessObject
            End If
        End With
    Next

    modImportExport.ExportMultipleObjects dAccessObjects, False

ExitProc:
    Exit Sub

ErrHandler:
    If DebugMode(True) Then
        Stop ' Use the unreachable Resume to return to the original line that caused error.
    End If
    Resume ExitProc
    Resume ' Use for debugging only
End Sub
