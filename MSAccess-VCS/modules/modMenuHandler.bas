Option Compare Database
Option Explicit

Private m_DefaultModel As New clsModelGitHub
Private m_Model As IVersionControl


'---------------------------------------------------------------------------------------
' Procedure : LoadVBEMenuForVCS
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Initialize the menu for VCS
'---------------------------------------------------------------------------------------
'
Public Sub LoadVBEMenuForVCS(Optional cModel As IVersionControl)
    
    If cModel Is Nothing Then Set cModel = DefaultModel
    Set m_Model = cModel
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DefaultModel
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Set up the default model
'---------------------------------------------------------------------------------------
'
Private Function DefaultModel() As IVersionControl

    Dim cGitHub As clsModelGitHub
    Dim strPath As String

    ' If we are editing the MSAccess-VCS project, then assume we are using GitHub
    ' Otherwise, use whatever is specified as the default model.
    If SelectionInActiveProject And VBE.ActiveVBProject.name = "MSAccess-VCS" Then
    
        ' Build path to source files. (Assuming default installation of GitHub)
        strPath = GetDocumentsFolder & "\GitHub\msaccess-vcs-integration\MSAccess-VCS\"
        If Dir(strPath, vbDirectory) <> "" Then
            ' Use this folder after verifying with user.
            If MsgBox("Use local GitHub folder?", vbQuestion + vbYesNo) = vbYes Then
                Set cGitHub = New clsModelGitHub
                Set DefaultModel = cGitHub
                DefaultModel.ExportBaseFolder = strPath
            End If
        Else
            ' Can't find the local GitHub project.
            Set DefaultModel = m_DefaultModel
        End If
    
    Else
        ' Use default model
        Set DefaultModel = m_DefaultModel
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDocumentsFolder
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Get "My Documents" folder.
'---------------------------------------------------------------------------------------
'
Private Function GetDocumentsFolder() As String
    Dim objShell As Object
    Set objShell = CreateObject("WScript.Shell")
    GetDocumentsFolder = objShell.SpecialFolders("MyDocuments")
    Set objShell = Nothing
End Function