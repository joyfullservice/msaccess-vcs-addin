Option Explicit
Option Compare Database

'Private m_DefaultModel As clsModelGitHub
Private m_Model As IVersionControl


'---------------------------------------------------------------------------------------
' Procedure : LoadVCSModel
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Load the VCS model directly
'---------------------------------------------------------------------------------------
'
Public Function LoadVCSModel(Optional cModel As IVersionControl)
    ReleaseObjectReferences
    If cModel Is Nothing Then Set cModel = DefaultModel
    Set m_Model = cModel
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadVersionControl
' Author    : Adam Waller
' Date      : 1/24/2018
' Purpose   : Load the version control system using the specified parameters
'           : (We use a collection to avoid the need for early binding on
'           :  class models in parent applications)
'---------------------------------------------------------------------------------------
'
Public Sub LoadVersionControl(colParams As Collection)

    Dim cModel As IVersionControl
    Dim varParam As Variant
    Dim strKey As String
    Dim strVal As String
    Dim strMsg As String
    
    ' Unload and clear any existing objects
    If Not m_Model Is Nothing Then
        m_Model.Terminate
        Set m_Model = Nothing
    End If
    
    ' Load parameters
    For Each varParam In colParams
        If IsArray(varParam) Then
            strKey = varParam(0)
            strVal = varParam(1)
            Select Case strKey
                Case "System"
                    Select Case strVal
                        Case "GitHub"
                            Set cModel = New clsModelGitHub
                        Case "SVN"
                            Set cModel = New clsModelSVN
                        Case Else
                            strMsg = "System not supported: " & strVal
                    End Select
                
                Case "Export Folder"
                    If Not cModel Is Nothing Then cModel.ExportBaseFolder = strVal
                    
                Case "Show Debug"
                    If Not cModel Is Nothing Then cModel.ShowDebug = strVal
                    
                Case "Include VBE"
                    If Not cModel Is Nothing Then cModel.IncludeVBE = strVal
                
                Case "Fast Save"
                    If Not cModel Is Nothing Then cModel.FastSave = strVal
                    
                Case "Save Table"
                    If Not cModel Is Nothing Then cModel.TablesToSaveData.Add strVal
                
                Case "Save Print Vars"
                    'if not cmodel is nothing then cmodel.
                
                Case Else
                    strMsg = "Unknown parameter: " & strKey
            
            End Select
        Else
            strMsg = "Parameter must be passed as an array."
        End If
        If strMsg <> "" Then Exit For
    Next varParam
    
    If strMsg = "" Then
        ' Set model
        Set m_Model = cModel
    Else
        ' Show message if errors were encountered
        MsgBox strMsg, vbExclamation, "Version Control"
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DefaultModel
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Set up the default model
'---------------------------------------------------------------------------------------
'
Private Function DefaultModel() As IVersionControl

    Dim strPath As String
    Dim cDefault As New clsModelGitHub

    ' If we are editing the MSAccess-VCS project, then assume we are using GitHub
    ' Otherwise, use whatever is specified as the default model.
    If SelectionInActiveProject And VBE.ActiveVBProject.Name = "MSAccess-VCS" Then
    
        ' Build path to source files. (Assuming default installation of GitHub)
        strPath = GetDocumentsFolder & "\GitHub\msaccess-vcs-integration\MSAccess-VCS\"
        If Dir(strPath, vbDirectory) <> "" Then
            ' Use this folder after verifying with user.
            If MsgBox("Use local GitHub folder?", vbQuestion + vbYesNo) = vbYes Then
                Set DefaultModel = New clsModelGitHub
                With DefaultModel
                    .ExportBaseFolder = strPath
                    .ShowDebug = False  ' Simple output messages
                End With
            End If
        Else
            ' Can't find the local GitHub project.
            Set DefaultModel = cDefault
        End If
    
    Else
        ' Use default model
        Set DefaultModel = cDefault
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


'---------------------------------------------------------------------------------------
' Procedure : ReleaseObjectReferences
' Author    : Adam Waller
' Date      : 6/2/2015
' Purpose   : Releases object references to allow unload of project
'---------------------------------------------------------------------------------------
'
Public Function ReleaseObjectReferences()
    If Not m_Model Is Nothing Then
        m_Model.Terminate
        Set m_Model = Nothing
    End If
    Set colVerifiedPaths = Nothing
End Function