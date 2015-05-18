Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Require the functions outlined in IVersionControl
' (Allows us to use different data models with the same
'  programming logic.)
Implements IVersionControl


' Exposed properties
Public ExportBaseFolder As String


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
    VCSSourcePath = Me.ExportBaseFolder
    If ProjectIsSelected Then
        ' Commit entire project
        ExportAllSource False
    Else
        ' Commit single file
        ExportByVBEComponent VBE.SelectedVBComponent
    End If
    ActivateGitHub
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
    strPath = Environ$("localappdata") & "\GitHub\GitHub.appref-ms"
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