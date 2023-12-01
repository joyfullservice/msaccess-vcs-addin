VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbRelation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This class extends the IDbComponent class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private m_Relation As DAO.Relation
Private m_Items(True To False) As Dictionary
Private m_FileList As Dictionary
Private m_Dbs As DAO.Database

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 4/30/2020
' Purpose   : Release reference to current db
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    Set m_Relation = Nothing
    Set m_Dbs = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export(Optional strAlternatePath As String)
    Dim strContent As String
    strContent = GetSource
    WriteFile strContent, Nz2(strAlternatePath, IDbComponent_SourceFile)
    VCSIndex.Update Me, IIf(strAlternatePath = vbNullString, eatExport, eatAltExport), GetStringHash(strContent, True)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim dItem As Dictionary
    Dim dFile As Dictionary
    Dim dField As Dictionary
    Dim fld As DAO.Field
    Dim rel As DAO.Relation

    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub

    ' Parse json file
    Set dFile = ReadJsonFile(strFile)
    If Not dFile Is Nothing Then

        ' Create new relation
        Set m_Dbs = CurrentDb
        Set dItem = dFile("Items")
        Set rel = m_Dbs.CreateRelation(dItem("Name"), dItem("Table"), dItem("ForeignTable"))
        rel.Attributes = dItem("Attributes")

        ' Add fields, and append to relation
        For Each dField In dItem("Fields")
            Set fld = rel.CreateField(dField("Name"))
            fld.ForeignName = dField("ForeignName")
            rel.Fields.Append fld
        Next dField

        ' Relationships create indexes, so we need to make sure an index
        ' with this name doesn't already exist. (Also check to be sure that
        ' we don't already have a relationship with this name.)
        LogUnhandledErrors
        On Error Resume Next
        With m_Dbs
            .TableDefs(rel.Table).Indexes.Delete rel.Name
            .TableDefs(rel.ForeignTable).Indexes.Delete rel.Name
            .Relations.Delete rel.Name
        End With
        CatchAny eelNoError, vbNullString, , False

        ' Add relationship to database
        m_Dbs.Relations.Append rel
        Set m_Relation = rel
    End If

    ' Update index
    VCSIndex.Update Me, eatImport, GetDictionaryHash(GetDictionary)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Merge
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Merge the source file into the existing database, updating or replacing
'           : any existing object.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Merge(strFile As String)
    ' Import if file exists
    If FSO.FileExists(strFile) Then
        ' Existing relationships will be removed during import
        IDbComponent_Import strFile
    Else
        ' If the source file was deleted, remove any existing relationship by this
        ' name during a merge operation.
        LogUnhandledErrors
        On Error Resume Next
        CurrentDb.Relations.Delete FSO.GetBaseName(strFile)
        VCSIndex.Remove Me, strFile
        If Err Then Err.Clear
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IDbComponent_MoveSource
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Move the component's source file(s) from one folder to another
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_MoveSource(strFromFolder As String, strToFolder As String)
    MoveFileIfExists strFromFolder & FSO.GetFileName(IDbComponent_SourceFile), strToFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary

    Dim rel As Relation
    Dim cRelation As IDbComponent
    Dim blnAdd As Boolean

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then

        ' Maintain persistent reference to database object so we don't
        ' lose the reference to the relation object with this procedure
        ' goes out of scope. (Make sure we release this on termination)
        If m_Dbs Is Nothing Then Set m_Dbs = CurrentDb

        Set m_Items(blnModifiedOnly) = New Dictionary
        For Each rel In m_Dbs.Relations
            ' Navigation pane groups are handled separately
            If Not (rel.Name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" _
                Or rel.Name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups" _
                Or IsInherited(rel)) Then
                Set cRelation = New clsDbRelation
                Set cRelation.DbObject = rel
                blnAdd = True
                If blnModifiedOnly Then blnAdd = cRelation.IsModified
                If blnAdd Then m_Items(blnModifiedOnly).Add cRelation.SourceFile, cRelation
            End If
        Next rel
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_Items(blnModifiedOnly)

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsInherited
' Author    : Adam Waller
' Date      : 6/30/2020
' Purpose   : Returns true if the relationship was inherited from tables in a linked
'           : database. (We don't need to export or import these.)
'---------------------------------------------------------------------------------------
'
Private Function IsInherited(objRelation As Relation) As Boolean
    IsInherited = ((objRelation.Attributes And dbRelationInherited) = dbRelationInherited)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetRelationFileName
' Author    : Adam Waller
' Date      : 6/4/2015
' Purpose   : Build file name based on relation name, including support for linked
'           : tables that would put a slash in the relation name.
'           : (Strips the link path from the table name)
'---------------------------------------------------------------------------------------
'
Public Function GetRelationFileName(objRelation As Relation) As String

    Dim strName As String

    strName = objRelation.Name

    If InStr(1, strName, "].") > 0 Then
        ' Need to remove path to linked file
        GetRelationFileName = GetSafeFileName(CStr(Split(strName, "].")(1)))
    Else
        GetRelationFileName = GetSafeFileName(strName)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSource
' Author    : Adam Waller
' Date      : 2/14/2022
' Purpose   : Return the full content that will be saved to the source file.
'---------------------------------------------------------------------------------------
'
Private Function GetSource() As String
    GetSource = BuildJsonFile(TypeName(Me), GetDictionary, "Database relationship")
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 5/6/2023
' Purpose   : Return a dictionary object of relationship definition.
'---------------------------------------------------------------------------------------
'
Private Function GetDictionary() As Dictionary

    Dim dItem As Dictionary
    Dim dField As Dictionary
    Dim colItems As Collection
    Dim fld As DAO.Field

    ' Relation properties
    Set dItem = New Dictionary
    With dItem
        .Add "Name", m_Relation.Name
        .Add "Attributes", m_Relation.Attributes
        .Add "Table", m_Relation.Table
        .Add "ForeignTable", m_Relation.ForeignTable
    End With

    ' Fields
    Set colItems = New Collection
    For Each fld In m_Relation.Fields
        Set dField = New Dictionary
        With dField
            .Add "Name", fld.Name
            .Add "ForeignName", fld.ForeignName
        End With
        colItems.Add dField
    Next fld
    dItem.Add "Fields", colItems

    ' Return dictionary object
    Set GetDictionary = dItem

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Dictionary
    If m_FileList Is Nothing Then Set m_FileList = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.json")
    Set IDbComponent_GetFileList = m_FileList
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_ClearOrphanedSourceFiles()
    ClearFilesByExtension IDbComponent_BaseFolder, "txt"
    ClearOrphanedSourceFiles Me, "json"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedDatabaseObjects
' Author    : Adam Waller
' Date      : 11/3/2021
' Purpose   : Remove database objects that are not represented by existing source files.
'---------------------------------------------------------------------------------------
'
Public Sub IDbComponent_ClearOrphanedDatabaseObjects()
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsModified
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Returns true if the object in the database has been modified since
'           : the last export of the object.
'---------------------------------------------------------------------------------------
'
Public Function IDbComponent_IsModified() As Boolean
    IDbComponent_IsModified = VCSIndex.Item(Me).FileHash <> GetStringHash(GetSource, True)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SourceIsModified
' Author    : Adam Waller
' Date      : 12/1/2023
' Purpose   : Returns true if the source file appears to be modified since the last
'           : import or export. (When compared to the index)
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_SourceIsModified(strFile As String) As Boolean

End Function


'---------------------------------------------------------------------------------------
' Procedure : DateModified
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'           : If the modified date cannot be determined (such as application
'           : properties) then this function will return 0.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_DateModified() As Date
    IDbComponent_DateModified = 0
End Function


'---------------------------------------------------------------------------------------
' Procedure : SourceModifiedDate
' Author    : Adam Waller
' Date      : 4/27/2020
' Purpose   : The date/time the source object was modified. In most cases, this would
'           : be the date/time of the source file, but it some cases like SQL objects
'           : the date can be determined through other means, so this function
'           : allows either approach to be taken.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_SourceModifiedDate() As Date
    IDbComponent_SourceModifiedDate = GetLastModifiedDate(IDbComponent_SourceFile)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "Relations"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "relations" & PathSep
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    If m_Relation Is Nothing Then Exit Property
    IDbComponent_Name = m_Relation.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    If m_Relation Is Nothing Then Exit Property
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetRelationFileName(m_Relation) & ".json"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count(Optional blnModifiedOnly As Boolean = False) As Long
    IDbComponent_Count = IDbComponent_GetAllFromDB(blnModifiedOnly).Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : QuickCount
' Author    : Adam Waller
' Date      : 6/14/2022
' Purpose   : Return a cached, non-iterative approximate count of database objects
'           : for use with progress indicators when scanning for changes. Single file
'           : objects like database properties can simply return 1.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_QuickCount() As Long
    Static lngCount As Long
    If lngCount = 0 Then lngCount = CurrentDb.Relations.Count - 1
    IDbComponent_QuickCount = lngCount + 1
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbRelation
End Property


'---------------------------------------------------------------------------------------
' Procedure : Upgrade
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Run any version specific upgrade processes before importing.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Upgrade()
    ' No upgrade needed.
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_Relation
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Relation = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Returns true if the export of all items is done as a single file instead
'           : of individual files for each component. (I.e. properties, references)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SingleFile() As Boolean
    IDbComponent_SingleFile = False
End Property


'---------------------------------------------------------------------------------------
' Procedure : Parent
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Return a reference to this class as an IDbComponent. This allows you
'           : to reference the public methods of the parent class without needing
'           : to create a new class object.
'---------------------------------------------------------------------------------------
'
Public Property Get Parent() As IDbComponent
    Set Parent = Me
End Property
