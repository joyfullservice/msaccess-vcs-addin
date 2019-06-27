Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Database

' Require the functions outlined in IVersionControl
' (Allows us to use different data models with the same
'  programming logic.)
Implements IVersionControl
Private m_vcs As IVersionControl

' Local instance of menu class
Private m_Menu As clsVbeMenu


'---------------------------------------------------------------------------------------
' Procedure : Terminate
' Author    : Adam Waller
' Date      : 6/2/2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function Terminate()
    Call Class_Terminate
End Function


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Initialize the class and load the menu.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set m_vcs = New IVersionControl
    Set m_Menu = New clsVbeMenu
    m_Menu.Construct Me
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 6/2/2015
' Purpose   : Remove reference to menu class
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    If Not m_Menu Is Nothing Then
        m_Menu.Terminate
        m_vcs.Terminate
        Set m_Menu = Nothing
        Set m_vcs = Nothing
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IVersionControl_Commit
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Activate GitHub. (do this through GitHub)
'---------------------------------------------------------------------------------------
'
Private Sub IVersionControl_Commit()
    ActivateGitHub
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IVersionControl_Diff
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Activate GitHub. (do this through GitHub)
'---------------------------------------------------------------------------------------
'
Private Sub IVersionControl_Diff()
    ActivateGitHub
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IVersionControl_Export
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Export the source code
'---------------------------------------------------------------------------------------
'
Private Sub IVersionControl_Export()
    If ProjectIsSelected Then
        ' Commit entire project
        ExportAllSource Me
    Else
        ' Commit single file
        ExportByVBEComponent VBE.SelectedVBComponent, Me
    End If
    ActivateGitHub
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IVersionControl_ExportAll
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Export all source
'---------------------------------------------------------------------------------------
'
Private Sub IVersionControl_ExportAll()
    ExportAllSource Me
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IVersionControl_HasRequiredSoftware
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Look for the application prefs file to verify installation.
'---------------------------------------------------------------------------------------
'
Private Property Get IVersionControl_HasRequiredSoftware(blnWarnUser As Boolean) As Boolean
    Dim strPath As String
    strPath = Environ$("localappdata") & "\GitHubDesktop\GitHubDesktop.exe" ' "\GitHub\GitHub.appref-ms"
    If Dir(strPath) <> "" Then
        ' Found path
        IVersionControl_HasRequiredSoftware = True
    Else
        If blnWarnUser Then MsgBox "Could not find GitHub Windows Client installation.", vbExclamation
    End If
End Property


'---------------------------------------------------------------------------------------
' Procedure : ActivateGitHub
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Activate the GitHub application
'---------------------------------------------------------------------------------------
'
Private Sub ActivateGitHub()
    On Error Resume Next
    AppActivate "GitHub"
    If Err Then
        Err.Clear
        MsgBox "GitHub application window not found. Is GitHub open?", vbExclamation
    End If
    On Error GoTo 0
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IVersionControl_Terminate
' Author    : Adam Waller
' Date      : 6/2/2015
' Purpose   : Terminate child classes
'---------------------------------------------------------------------------------------
'
Private Sub IVersionControl_Terminate()
    Call Class_Terminate
End Sub


'---------------------------------------------------------------------------------------
'///////////////////////////////////////////////////////////////////////////////////////
'---------------------------------------------------------------------------------------
' Procedure : (Multiple)
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Wrapper classes to call functions in parent class
'---------------------------------------------------------------------------------------
'
Private Property Get IVersionControl_TablesToSaveData() As Collection
    Set IVersionControl_TablesToSaveData = m_vcs.TablesToSaveData
End Property
Private Property Let IVersionControl_ExportBaseFolder(ByVal RHS As String)
    m_vcs.ExportBaseFolder = RHS
End Property
Private Property Get IVersionControl_ExportBaseFolder() As String
    IVersionControl_ExportBaseFolder = m_vcs.ExportBaseFolder
End Property
Private Property Let IVersionControl_ShowDebug(ByVal RHS As Boolean)
    m_vcs.ShowDebug = RHS
End Property
Private Property Get IVersionControl_ShowDebug() As Boolean
    IVersionControl_ShowDebug = m_vcs.ShowDebug
End Property
Private Property Let IVersionControl_IncludeVBE(ByVal RHS As Boolean)
    m_vcs.IncludeVBE = RHS
End Property
Private Property Get IVersionControl_IncludeVBE() As Boolean
    IVersionControl_IncludeVBE = m_vcs.IncludeVBE
End Property
Private Property Get IVersionControl_SelectionSourceFile(Optional UseVBEFile As Boolean = True) As String
    IVersionControl_SelectionSourceFile = m_vcs.SelectionSourceFile(UseVBEFile)
End Property
Private Property Let IVersionControl_FastSave(ByVal RHS As Boolean)
    m_vcs.FastSave = RHS
End Property
Private Property Get IVersionControl_FastSave() As Boolean
    IVersionControl_FastSave = m_vcs.FastSave
End Property
Private Property Let IVersionControl_SavePrintVars(ByVal RHS As Boolean)
    m_vcs.SavePrintVars = RHS
End Property
Private Property Get IVersionControl_SavePrintVars() As Boolean
    IVersionControl_SavePrintVars = m_vcs.SavePrintVars
End Property
Private Sub IVersionControl_Log(strText As String, Optional blnPrint As Boolean = True, Optional blnStartNewLine As Boolean = True)
    m_vcs.Log strText, blnPrint, blnStartNewLine
End Sub
Private Sub IVersionControl_SaveLogFile(strPath As String)
    m_vcs.SaveLogFile strPath
End Sub
Private Property Let IVersionControl_StripPublishOption(ByVal RHS As Boolean)
    m_vcs.StripPublishOption = RHS
End Property
Private Property Get IVersionControl_StripPublishOption() As Boolean
    IVersionControl_StripPublishOption = m_vcs.StripPublishOption
End Property
Private Property Let IVersionControl_AggressiveSanitize(ByVal RHS As Boolean)
    m_vcs.AggressiveSanitize = RHS
End Property
Private Property Get IVersionControl_AggressiveSanitize() As Boolean
    IVersionControl_AggressiveSanitize = m_vcs.AggressiveSanitize
End Property
Private Property Let IVersionControl_SaveQuerySQL(ByVal RHS As Boolean)
    m_vcs.SaveQuerySQL = RHS
End Property
Private Property Get IVersionControl_SaveQuerySQL() As Boolean
    IVersionControl_SaveQuerySQL = m_vcs.SaveQuerySQL
End Property
Private Property Let IVersionControl_SaveTableSQL(ByVal RHS As Boolean)
    m_vcs.SaveTableSQL = RHS
End Property
Private Property Get IVersionControl_SaveTableSQL() As Boolean
    IVersionControl_SaveTableSQL = m_vcs.SaveTableSQL
End Property