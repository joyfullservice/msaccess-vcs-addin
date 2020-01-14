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
Public Function LoadVersionControlMenu(colParams As Collection) As IVersionControl

    Dim cModel As IVersionControl
    Dim varParam As Variant
    Dim strKey As String
    Dim strVal As String
    Dim strMsg As String
    Dim strCurrent As String
    
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
                    If Not cModel Is Nothing Then cModel.SavePrintVars = strVal
                
                Case "Save Query SQL"
                    If Not cModel Is Nothing Then cModel.SaveQuerySQL = strVal
                
                Case "Save Table SQL"
                    If Not cModel Is Nothing Then cModel.SaveTableSQL = strVal
                
                Case Else
                    strMsg = "Unknown parameter: " & strKey
            
            End Select
        Else
            strMsg = "Parameter must be passed as an array."
        End If
        If strMsg <> "" Then Exit For
    Next varParam
    
    If strMsg = "" Then
            
        ' Make sure version matches to enable fast save.
        If Not cModel Is Nothing Then
            If cModel.FastSave Then
                strCurrent = GetVCSVersion
                If strCurrent <> "" And strCurrent = GetDBProperty("Last VCS Version") Then
                    ' Only allow fast save if we have run a full export with this
                    ' version of VCS.
                    cModel.FastSave = True
                Else
                    ' Require a full export on current version before enabling fast save.
                    cModel.FastSave = False
                End If
            End If
        End If
        
        ' Set model for class
        Set m_Model = cModel
        Set LoadVersionControlMenu = cModel
        
    Else
        ' Show message if errors were encountered
        MsgBox strMsg, vbExclamation, "Version Control"
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetVCSVersion
' Author    : Adam Waller
' Date      : 1/28/2019
' Purpose   : Gets the version of the version control system. (Used to turn off fast
'           : save until a full export has been run with the current version of
'           : the MSAccessVCS addin.)
'---------------------------------------------------------------------------------------
'
Public Function GetVCSVersion() As String
    
    Dim dbs As Database
    Dim objParent As Object
    Dim prp As Object
    
    Set objParent = CodeDb
    If objParent Is Nothing Then Set objParent = CurrentProject ' ADP support

    For Each prp In objParent.Properties
        If prp.Name = "AppVersion" Then
            ' Return version
            GetVCSVersion = prp.Value
        End If
    Next prp

End Function


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