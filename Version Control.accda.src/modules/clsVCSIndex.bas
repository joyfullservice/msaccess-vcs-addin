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

' General properties
Public MergeBuildDate As Date
Public FullBuildDate As Date
Public ExportDate As Date
Public FullExportDate As Date

' Git integration
Public LastMergedCommit As String

' Action types for update function
Public Enum eIndexActionType
    eatExport
    eatImport
End Enum

' Index of component/file information, based on source files
Private m_dIndex As Dictionary


'---------------------------------------------------------------------------------------
' Procedure : LoadFromFile
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Load the state for the project.
'---------------------------------------------------------------------------------------
'
Public Sub LoadFromFile()

    Dim dFile As Dictionary
    Dim dItem As Dictionary
    Dim varKey As Variant
    
    ' Reset class to blank values
    Call Class_Initialize
    
    ' Load properties
    If FSO.FileExists(FileName) Then
        Set dFile = ReadJsonFile(FileName)
        If Not dFile Is Nothing Then
            If dFile.Exists("Items") Then
                ' Load properties from class
                For Each varKey In dFile("Items").Keys
                    If m_dIndex.Exists(varKey) Then
                        If varKey = "Components" Then
                            Set dItem = dFile("Items")(varKey)
                            Set m_dIndex(varKey) = dItem
                        Else
                            ' Set property by name
                            CallByName Me, CStr(varKey), VbLet, Nz(dFile("Items")(varKey), 0)
                        End If
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

    Dim varCat As Variant
    Dim varKey As Variant
    Dim varValue As Variant
    Dim dComponents As Dictionary
    
    ' Load dictionary from properties
    For Each varKey In m_dIndex.Keys
        If varKey <> "Components" Then
            varValue = CallByName(Me, CStr(varKey), VbGet)
            ' Save blank dates as null
            If Right(varKey, 4) = "Date" And varValue = 0 Then varValue = Null
            m_dIndex(varKey) = varValue
        End If
    Next varKey

    ' Sort files and components
    Set dComponents = m_dIndex("Components")
    For Each varCat In dComponents.Keys
        Set dComponents(varCat) = SortDictionaryByKeys(dComponents(varCat))
    Next varCat
    Set m_dIndex("Components") = SortDictionaryByKeys(dComponents)

    ' Save index to file
    WriteJsonFile Me, m_dIndex, FileName, "Version Control System Index"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Update
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Updates an item in the index.
'---------------------------------------------------------------------------------------
'
Public Function Update(cItem As IDbComponent, intAction As eIndexActionType, _
    Optional strHash As String, Optional dteDateTime As Date) As Dictionary
    
    Dim dItem As Dictionary
    
    ' Look up dictionary item, creating if needed.
    Set dItem = Me.GetItem(cItem)
    
    ' Update dictionary values
    With dItem
    
        ' Update hash
        If strHash = vbNullString Then
            ' Remove hash if not used.
            If .Exists("Hash") Then .Remove "Hash"
        Else
            .Item("Hash") = strHash
        End If
        
        ' Add timestamp, defaulting to now
        If dteDateTime = 0 Then dteDateTime = Now
        Select Case intAction
            Case eatExport: .Item("ExportDate") = dteDateTime
            Case eatImport: .Item("ImportDate") = dteDateTime
        End Select
    
    End With
    
    ' Return dictionary object in case the caller wants to
    ' manipulate additional values.
    Set Update = dItem
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetItem
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Returns a dictionary object with the saved values, creating if needed.
'---------------------------------------------------------------------------------------
'
Public Function GetItem(cItem As IDbComponent) As Dictionary
    
    Dim strFile As String
    
    ' Get just the file name from the path.
    strFile = FSO.GetFileName(cItem.SourceFile)
    
    ' Get or creat dictionary objects.
    With m_dIndex("Components")
        If Not .Exists(cItem.Category) Then Set .Item(cItem.Category) = New Dictionary
        If Not .Item(cItem.Category).Exists(strFile) Then Set .Item(cItem.Category)(strFile) = New Dictionary
        Set GetItem = .Item(cItem.Category)(strFile)
    End With
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : FileName
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Return file name for git state json file.
'---------------------------------------------------------------------------------------
'
Private Function FileName() As String
    FileName = Options.GetExportFolder & "vcs-index.json"
End Function


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Set up the dictionary object and keys for reflection.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    Set m_dIndex = New Dictionary
    
    With m_dIndex
        .Add "MergeBuildDate", Null
        .Add "FullBuildDate", Null
        .Add "ExportDate", Null
        .Add "FullExportDate", Null
        .Add "LastMergedCommit", vbNullString
        Set .Item("Components") = New Dictionary
    End With
    
End Sub