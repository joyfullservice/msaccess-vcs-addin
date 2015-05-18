Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Update these values to set the paths and commands for default installations
'---------------------------------------------------------------------------------------
Const cAppPath As String = "TortoiseSVN\bin\TortoiseProc.exe"
Const cCmdCommit As String = " /command:commit /notempfile /path:"
Const cCmdUpdate As String = " /command:update /rev /notempfile /path:"
' Differ
Const cDiffPath As String = "WinMerge\WinMergeU.exe"
Const cCmdDiff As String = ""
'---------------------------------------------------------------------------------------


' Require the functions outlined in IVersionControl
' (Allows us to use different data models with the same
'  programming logic.)
Implements IVersionControl

' Private variables
Private m_ProgFiles As String


'---------------------------------------------------------------------------------------
' Procedure : IVersionControl_Commit
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Commit to SVN
'---------------------------------------------------------------------------------------
'
Private Sub IVersionControl_Commit()
    If ProjectIsSelected Then
        ' Commit entire project
        ExportAllSource False
    Else
        ' Commit single file
        ExportByVBEComponent VBE.SelectedVBComponent
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IVersionControl_Diff
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Diff the file
'---------------------------------------------------------------------------------------
'
Private Sub IVersionControl_Diff()
    MsgBox "Diff"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : AppPath
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Wrapper for code readability
'---------------------------------------------------------------------------------------
'
Private Function AppPath() As String
    AppPath = ProgramFilesFolder & cAppPath
End Function



'---------------------------------------------------------------------------------------
' Procedure : IVersionControl_HasRequiredSoftware
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns true if the required program files are found
'---------------------------------------------------------------------------------------
'
Private Property Get IVersionControl_HasRequiredSoftware(blnWarnUser As Boolean) As Boolean
    Dim blnFound As Boolean
    If Dir(cAppPath) <> "" Then
        If Dir(cDiffPath) <> "" Then
            IVersionControl_HasRequiredSoftware = True
        Else
            
        End If
    Else
    
    End If
End Property
