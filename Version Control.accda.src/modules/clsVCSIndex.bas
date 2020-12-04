Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsGitIndex
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

' Reference to Git Integration
Private Git As clsGitIntegration

' Index of component/file information, based on source files
Private m_dIndex As Dictionary
Private m_strFile As String


'---------------------------------------------------------------------------------------
' Procedure : LoadFromFile
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Load the state for the project.
'---------------------------------------------------------------------------------------
'
Public Sub LoadFromFile(Optional strFile As String)

    Dim dFile As Dictionary
    Dim dItem As Dictionary
    Dim varKey As Variant
    
    ' Reset class to blank values
    Call Class_Initialize
    
    ' Load properties
    m_strFile = strFile
    If m_strFile = vbNullString Then m_strFile = DefaultFileName
    If FSO.FileExists(m_strFile) Then
        Set dFile = ReadJsonFile(m_strFile)
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
            If Right(varKey, 4) = "Date" Then
                m_dIndex(varKey) = ZNDate(CStr(varValue))
            Else
                m_dIndex(varKey) = CStr(varValue)
            End If
        End If
    Next varKey

    ' Sort files and components
    Set dComponents = m_dIndex("Components")
    For Each varCat In dComponents.Keys
        Set dComponents(varCat) = SortDictionaryByKeys(dComponents(varCat))
    Next varCat
    Set m_dIndex("Components") = SortDictionaryByKeys(dComponents)

    ' Save index to file
    If m_strFile <> vbNullString Then
        WriteJsonFile Me, m_dIndex, m_strFile, "Version Control System Index"
    End If
    
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
    Set dItem = Me.Item(cItem)
    
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
            Case eatExport: .Item("ExportDate") = CStr(dteDateTime)
            Case eatImport: .Item("ImportDate") = CStr(dteDateTime)
        End Select
        
        ' Save timestamp of exported source file.
        dteDateTime = GetLastModifiedDate(cItem.SourceFile)
        .Item("SourceModified") = ZNDate(CStr(dteDateTime))
    
    End With
    
    ' Return dictionary object in case the caller wants to
    ' manipulate additional values.
    Set Update = dItem
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : Remove
' Author    : Adam Waller
' Date      : 12/2/2020
' Purpose   : Remove an item from the index when the object and file no longer exist.
'---------------------------------------------------------------------------------------
'
Public Sub Remove(cItem As IDbComponent)
    
    Dim strFile As String
    
    ' Get just the file name from the path.
    strFile = FSO.GetFileName(cItem.SourceFile)
    
    ' Remove dictionary objects.
    With m_dIndex("Components")
        If .Exists(cItem.Category) Then
            If .Item(cItem.Category).Exists(strFile) Then
                .Item(cItem.Category).Remove strFile
                ' Remove category if no more items
                If .Item(cItem.Category).Count = 0 Then
                    .Remove cItem.Category
                End If
            End If
        End If
    End With
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetItem
' Author    : Adam Waller
' Date      : 11/30/2020
' Purpose   : Returns a dictionary object with the saved values, creating if needed.
'---------------------------------------------------------------------------------------
'
Public Function Item(cItem As IDbComponent) As Dictionary
    
    Dim strFile As String
    
    ' Get just the file name from the path.
    strFile = FSO.GetFileName(cItem.SourceFile)
    
    ' Get or create dictionary objects.
    With m_dIndex("Components")
        If Not .Exists(cItem.Category) Then Set .Item(cItem.Category) = New Dictionary
        If Not .Item(cItem.Category).Exists(strFile) Then Set .Item(cItem.Category)(strFile) = New Dictionary
        Set Item = .Item(cItem.Category)(strFile)
    End With
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : Exists
' Author    : Adam Waller
' Date      : 12/2/2020
' Purpose   : Returns true if the file exists in the index.
'---------------------------------------------------------------------------------------
'
Public Function Exists(cCategory As IDbComponent, strSourceFilePath As String) As Boolean

    Dim strFile As String
    Dim blnExists
    
    ' Get just the file name from the path.
    strFile = FSO.GetFileName(strSourceFilePath)
    
    ' See if the entry exists in the index
    With m_dIndex("Components")
        If .Exists(cCategory.Category) Then
            blnExists = .Item(cCategory.Category).Exists(strFile)
        End If
    End With
    
    ' Return result
    Exists = blnExists
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetModifiedSourceFiles
' Author    : Adam Waller
' Date      : 12/2/2020
' Purpose   : Return a collection of source files that appear to be modified from
'           : the previous export. If Git integration is enabled, this will be used
'           : to improve the performance of the determination. Otherwise file modified
'           : dates will be used to determine which files have changed.
'           : NOTE: This will also include paths for files that no longer exist
'           : in source files.
'---------------------------------------------------------------------------------------
'
Public Function GetModifiedSourceFiles(cCategory As IDbComponent) As Collection

    Dim colAllFiles As Collection
    Dim varFile As Variant
    Dim strFile As String
    Dim strPath As String
    Dim blnModified As Boolean
    
    ' Get a list of all the files for this component.
    Set colAllFiles = cCategory.GetFileList

    Set GetModifiedSourceFiles = New Collection
    With GetModifiedSourceFiles
        For Each varFile In colAllFiles
            strFile = varFile
            ' Reset flag
            blnModified = True
            If Me.Exists(cCategory, strFile) Then
                strPath = Join(Array("Components", cCategory.Category, FSO.GetFileName(strFile), "SourceModified"), "\")
                ' Compare modified date of file with modified date in index.
                blnModified = Not dNZ(m_dIndex, strPath) = GetLastModifiedDate(strFile)
            End If
            ' Add modified files to collection
            If blnModified Then .Add strFile
        Next varFile
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : FileName
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Return file name for git state json file.
'---------------------------------------------------------------------------------------
'
Private Function DefaultFileName() As String
    DefaultFileName = Options.GetExportFolder & "vcs-index.json"
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
    
    ' Load Git integration
    Set Git = New clsGitIntegration
    
End Sub