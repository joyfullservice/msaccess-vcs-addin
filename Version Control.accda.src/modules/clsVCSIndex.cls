VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVCSIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsVCSIndex
' Author    : Adam Waller
' Date      : 11/25/2020
' Purpose   : Maintain index of source files and database objects so that changes
'           : can be detected.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

' File name for index
Private Const cstrFileName As String = "vcs-index.json"

' General properties
Public MergeBuildDate As Date
Public FullBuildDate As Date
Public ExportDate As Date
Public FullExportDate As Date
Public OptionsHash As String
Public Disabled As Boolean

' Git integration
Public LastMergedCommit As String

' Action types for update function
Public Enum eIndexActionType
    eatExport
    eatImport
End Enum

' Index of component/file information, based on source files
Private m_dIndex As Dictionary
Private m_dGitIndex As Dictionary
Private m_strFile As String
Private m_Git As clsGitIntegration
Private m_Conflicts As clsConflicts


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

    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Sub

    ' Reset class to blank values
    Call Class_Initialize
        
    ' Load properties
    m_strFile = strFile
    If m_strFile = vbNullString Then m_strFile = DefaultFilePath
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
Public Sub Save(Optional strInFolder As String)

    Dim varCat As Variant
    Dim varKey As Variant
    Dim varValue As Variant
    Dim dComponents As Dictionary
    Dim strFile As String

    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Sub

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

    ' Build file path
    If strInFolder = vbNullString Then
        strFile = m_strFile
    Else
        strFile = StripSlash(strInFolder) & PathSep & cstrFileName
    End If

    ' Save index to file
    If m_strFile <> vbNullString Then
        WriteJsonFile TypeName(Me), m_dIndex, strFile, "Version Control System Index"
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
    
    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Function
    
    ' Look up dictionary item, creating if needed.
    Set dItem = Me.Item(cItem).dParent
    
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
        
        ' Save hash of file properties
        .Item("FilePropertiesHash") = GetFilePropertyHash(cItem.SourceFile)
    
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

    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Sub

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
Public Function Item(cItem As IDbComponent, Optional strSourceFile As String) As clsVCSIndexItem

    Dim strFile As String
    
    ' Exit if we have disabled the index functionality
    If Me.Disabled Then Exit Function

    ' Get just the file name from the path.
    strFile = FSO.GetFileName(Nz2(strSourceFile, cItem.SourceFile))
    Set Item = LoadItem(cItem, strFile)
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadItem
' Author    : Adam Waller
' Date      : 9/23/2021
' Purpose   : Loads the item from the index, creating if it does not exist.
'---------------------------------------------------------------------------------------
'
Private Function LoadItem(cItem As IDbComponent, strSourceFile As String) As clsVCSIndexItem

    Dim dItem As Dictionary
    
    ' Get or create dictionary objects.
    With m_dIndex("Components")
        If Not .Exists(cItem.Category) Then Set .Item(cItem.Category) = New Dictionary
        If Not .Item(cItem.Category).Exists(strSourceFile) Then Set .Item(cItem.Category)(strSourceFile) = New Dictionary
        Set dItem = .Item(cItem.Category)(strSourceFile)
        Set LoadItem = New clsVCSIndexItem
        With LoadItem
            ' Load properties to class
            Set .dParent = dItem
            .ExportDate = Nz2(dNZ(dItem, "ExportDate"), 0)
            .ImportDate = Nz2(dNZ(dItem, "ImportDate"), 0)
            .SourceModified = Nz2(dNZ(dItem, "SourceModified"), 0)
            .FilePropertiesHash = dNZ(dItem, "FilePropertiesHash")
            .Hash = dNZ(dItem, "Hash")
        End With
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
' Procedure : DefaultDevModeHash
' Author    : Adam Waller
' Date      : 2/17/2021
' Purpose   : Return a hash of the printer settings from the default printer. (Used to
'           : determine whether a form or report is using any specific or custom
'           : print settings, thereby requiring us to save the print settings.
'---------------------------------------------------------------------------------------
'
Public Property Get DefaultDevModeHash() As String

    ' Cache the result for future calls
    Static strHash As String
    
    If strHash = vbNullString Then
        With New clsDevMode
            .LoadFromDefaultPrinter
            strHash = .GetHash
        End With
    End If
    
    ' Return hash
    DefaultDevModeHash = strHash
    
End Property


'---------------------------------------------------------------------------------------
' Procedure : GetFilePropertyHash
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Returns a hash of some file properties used to quickly scan for changes.
'---------------------------------------------------------------------------------------
'
Public Function GetFilePropertyHash(strFile As String) As String

    Dim oFile As Scripting.File
    
    If FSO.FileExists(strFile) Then
        Perf.OperationStart "Get File Property Hash"
        Set oFile = FSO.GetFile(strFile)
        
        With New clsConcat
            .Add oFile.DateLastModified, oFile.Size
            GetFilePropertyHash = GetStringHash(.GetStr)
        End With
        Perf.OperationEnd
    End If

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
    Dim dItem As Dictionary
    Dim blnModified As Boolean
    
    Set GetModifiedSourceFiles = New Collection
    With GetModifiedSourceFiles
        ' Make sure we have a previous commit to reference.
        If Options.UseGitIntegration And LastMergedCommit <> vbNullString Then
            ' Return list of any modified files of this type as served by Git.
            Set GetModifiedSourceFiles = Git.GetModifiedSourceFiles(cCategory)
        Else
            ' Loop through files and check modified dates.
            ' Get a list of all the files for this component.
            Set colAllFiles = cCategory.GetFileList
            For Each varFile In colAllFiles
                strFile = varFile
                ' Reset flag
                blnModified = True
                If Me.Exists(cCategory, strFile) Then
                    ' Build the path to find the item in the index.
                    strPath = Join(Array("Components", cCategory.Category, FSO.GetFileName(strFile), "SourceModified"), PathSep)
                    ' Compare modified date of file with modified date in index.
                    ' File is considered not modified if the index date matches the file modification date.
                    blnModified = Not dNZ(m_dIndex, strPath) = GetLastModifiedDate(strFile)
                End If
                ' Add modified files to collection
                If blnModified Then .Add strFile
            Next varFile
        End If
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : Conflicts
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Expose conflicts as an extension of the index
'---------------------------------------------------------------------------------------
'
Public Function Conflicts() As clsConflicts
    If m_Conflicts Is Nothing Then Set m_Conflicts = New clsConflicts
    Set Conflicts = m_Conflicts
End Function


Public Sub CheckExportConflicts(cComponents As Collection)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckImportConflicts
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Check for potential conflicts when merging this object.
'---------------------------------------------------------------------------------------
'
Public Sub CheckImportConflicts(cComponent As IDbComponent, dFiles As Dictionary)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Git
' Author    : Adam Waller
' Date      : 1/19/2021
' Purpose   : Reference to the Git integration class
'---------------------------------------------------------------------------------------
'
Private Function Git() As clsGitIntegration
    If m_Git Is Nothing Then Set m_Git = New clsGitIntegration
    Set Git = m_Git
End Function


'---------------------------------------------------------------------------------------
' Procedure : FileName
' Author    : Adam Waller
' Date      : 11/24/2020
' Purpose   : Return file name for git state json file.
'---------------------------------------------------------------------------------------
'
Private Function DefaultFilePath() As String
    If DatabaseOpen Then DefaultFilePath = Options.GetExportFolder & cstrFileName
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
        .Add "OptionsHash", vbNullString
        .Add "LastMergedCommit", vbNullString
        Set .Item("Components") = New Dictionary
    End With
    
    ' Load Git integration
    Set m_dGitIndex = Nothing
    
End Sub



