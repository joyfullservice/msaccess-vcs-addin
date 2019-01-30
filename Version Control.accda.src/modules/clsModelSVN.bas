Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Database

'---------------------------------------------------------------------------------------
' Update these values to set the paths and commands for default installations
'---------------------------------------------------------------------------------------
Const cAppPath As String = "TortoiseSVN\bin\TortoiseProc.exe"
Const cCmdCommit As String = " /command:commit /path:"
Const cCmdUpdate As String = " /command:update /path:"
Const cCmdDiff As String = " /command:diff /path:"
'---------------------------------------------------------------------------------------


' Require the functions outlined in IVersionControl
' (Allows us to use different data models with the same
'  programming logic.)
Implements IVersionControl
Private m_vcs As IVersionControl

' Private variables
Private m_Menu As clsVbeMenu


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
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 6/2/2015
' Purpose   : Remove reference to menu class
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    If Not m_Menu Is Nothing Then
        m_Menu.Terminate
        Set m_Menu = Nothing
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IVersionControl_Commit
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Commit to SVN
'---------------------------------------------------------------------------------------
'
Private Sub IVersionControl_Commit()
    Call IVersionControl_Export
    ' For some reason we have issues when we try to use the VBA Shell command
    ' The VBScript version seems to work fine.
    Shell2 AppPath & cCmdCommit & m_vcs.SelectionSourceFile
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IVersionControl_Diff
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Diff the file
'---------------------------------------------------------------------------------------
'
Private Sub IVersionControl_Diff()
    Call IVersionControl_Export
    Shell AppPath & cCmdDiff & m_vcs.SelectionSourceFile
End Sub


'---------------------------------------------------------------------------------------
' Procedure : AppPath
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Wrapper for code readability
'---------------------------------------------------------------------------------------
'
Private Function AppPath() As String
    Dim strPath As String
    strPath = ProgramFilesFolder
    ' Assume we are using the 64-bit version
    strPath = Replace(strPath, " (x86)", "")
    AppPath = """" & strPath & cAppPath & """"
End Function


'---------------------------------------------------------------------------------------
' Procedure : IVersionControl_Export
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Export files for SVN
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
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IVersionControl_ExportAll
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Export all source files for the project.
'---------------------------------------------------------------------------------------
'
Private Sub IVersionControl_ExportAll()
    ExportAllSource Me
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IVersionControl_HasRequiredSoftware
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns true if the required program files are found
'---------------------------------------------------------------------------------------
'
Private Property Get IVersionControl_HasRequiredSoftware(blnWarnUser As Boolean) As Boolean
    Dim strMsg As String
    Dim strPath As String
    strPath = Replace(AppPath, """", "")
    If Dir(strPath) <> "" Then
        IVersionControl_HasRequiredSoftware = True
    Else
        strMsg = "Could not find SVN program in " & vbCrLf & strPath
    End If
    If strMsg <> "" And blnWarnUser Then MsgBox strMsg, vbExclamation, "Version Control"
End Property


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