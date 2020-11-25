Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsGitState
' Author    : Adam Waller
' Date      : 11/25/2020
' Purpose   :
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public LastMergedCommit As String
Public MergeBuildTime As Date


Private m_dState As Dictionary


'---------------------------------------------------------------------------------------
' Procedure : LoadFromFile
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Load the state for the project.
'---------------------------------------------------------------------------------------
'
Public Sub LoadFromFile()

    Dim dFile As Dictionary
    Dim varKey As Variant
    
    ' Reset class to blank values
    Call Class_Initialize
    
    If FSO.FileExists(FileName) Then
        Set dFile = ReadJsonFile(FileName)
        If Not dFile Is Nothing Then
            If dFile.Exists("Items") Then
                ' Load properties from class
                For Each varKey In dFile("Items")
                    If m_dState.Exists(varKey) Then
                        ' Set property by name
                        CallByName Me, CStr(varKey), VbLet, dFile("Items")(varKey)
                    End If
                Next varKey
            End If
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Save
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Save to a file
'---------------------------------------------------------------------------------------
'
Public Sub Save()

    Dim varKey As Variant
    
    ' Load dictionary from properties
    For Each varKey In m_dState.Keys
        m_dState(varKey) = CallByName(Me, CStr(varKey), VbGet)
    Next varKey

    WriteJsonFile Me, m_dState, FileName, "Git Integration Status"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetModifiedSourceFiles
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Return the modified source file paths for this component type.
'---------------------------------------------------------------------------------------
'
Public Function GetModifiedSourceFiles(cCat As IDbComponent) As Collection

End Function


'---------------------------------------------------------------------------------------
' Procedure : FileName
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Return file name for git state json file.
'---------------------------------------------------------------------------------------
'
Private Function FileName() As String
    FileName = Options.GetExportFolder & "git-sync.json"
End Function


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Set up the dictionary object and keys for reflection.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    Set m_dState = New Dictionary
    With m_dState
        .Add "LastMergedCommit", vbNullString
        .Add "MergeBuildTime", 0
    End With
End Sub