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
Private m_AllItems As Collection
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
    Set m_Dbs = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export()
    
    Dim dItem As Scripting.Dictionary
    Dim dField As Scripting.Dictionary
    Dim colItems As Collection
    Dim fld As DAO.Field
    
    ' Relation properties
    Set dItem = New Scripting.Dictionary
    With dItem
        .Add "Name", m_Relation.Name
        .Add "Attributes", m_Relation.Attributes
        .Add "Table", m_Relation.Table
        .Add "ForeignTable", m_Relation.ForeignTable
        .Add "PartialReplica", m_Relation.PartialReplica
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
    
    ' Write to json file
    WriteJsonFile Me, dItem, IDbComponent_SourceFile, "Database relationship"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)
    
    Dim dItem As Scripting.Dictionary
    Dim dFile As Dictionary
    Dim dField As Scripting.Dictionary
    Dim fld As DAO.Field
    Dim dbs As DAO.Database
    Dim rel As DAO.Relation
    
    ' Parse json file
    Set dFile = ReadJsonFile(strFile)
    If Not dFile Is Nothing Then
        
        ' Create new relation
        Set dbs = CurrentDb
        Set dItem = dFile("Items")
        Set rel = dbs.CreateRelation(dItem("Name"), dItem("Table"), dItem("ForeignTable"))
        rel.Attributes = dItem("Attributes")
        rel.PartialReplica = dItem("PartialReplica")
        
        ' Add fields, and append to relation
        For Each dField In dFile("Fields")
            Set fld = rel.CreateField(dField("Name"))
            fld.ForeignName = dField("ForeignName")
        Next dField
        rel.Fields.Append fld
        
        ' Add relationship to database
        dbs.Relations.Append rel
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB() As Collection
    
    Dim rel As Relation
    Dim cRelation As IDbComponent

    ' Build collection if not already cached
    If m_AllItems Is Nothing Then
    
        ' Maintain persistent reference to database object so we don't
        ' lose the reference to the relation object with this procedure
        ' goes out of scope. (Make sure we release this on termination)
        Set m_Dbs = CurrentDb
        
        Set m_AllItems = New Collection
        For Each rel In m_Dbs.Relations
            ' Navigation pane groups are handled separately
            If Not (rel.Name = "MSysNavPaneGroupsMSysNavPaneGroupToObjects" _
                Or rel.Name = "MSysNavPaneGroupCategoriesMSysNavPaneGroups") Then
                Set cRelation = New clsDbRelation
                Set cRelation.DbObject = rel
                m_AllItems.Add cRelation, rel.Name
            End If
        Next rel
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_AllItems
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportRelation
' Author    : Adam Kauffman
' Date      : 02/18/2020
' Purpose   : Import a table relationship
'---------------------------------------------------------------------------------------
'
Private Sub ImportRelation(ByRef filePath As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim thisDb As Database
    Set thisDb = appInstance.CurrentDb
    
    Dim fileLines() As String
    With FSO.OpenTextFile(filePath, IOMode:=ForReading, Create:=False, Format:=TristateFalse)
        fileLines = Split(.ReadAll, vbCrLf)
        .Close
    End With
    
    Dim newRelation As Relation
    Set newRelation = thisDb.CreateRelation(fileLines(1), fileLines(2), fileLines(3), fileLines(0))
    
    Dim newField As Field
    Dim thisLine As Long
    For thisLine = 4 To UBound(fileLines)
        If "Field = Begin" = fileLines(thisLine) Then
            thisLine = thisLine + 1
            Set newField = newRelation.CreateField(fileLines(thisLine))  ' Name set here
            thisLine = thisLine + 1
            newField.ForeignName = fileLines(thisLine)
            thisLine = thisLine + 1
            If "End" <> fileLines(thisLine) Then
                Set newField = Nothing
                Err.Raise 40000, "ImportRelation", "Missing 'End' for a 'Begin' in " & filePath
            End If
            
            newRelation.Fields.Append newField
        End If
    Next thisLine
        
    ' Remove conflicting Index entries because adding the relation creates new indexes causing "Error 3284 Index already exists"
    On Error Resume Next
    With thisDb
        .Relations.Delete newRelation.Name  ' Avoid 3012 Relationship already exists
        .TableDefs(newRelation.Table).Indexes.Delete newRelation.Name
        .TableDefs(newRelation.ForeignTable).Indexes.Delete newRelation.Name
    End With
    On Error GoTo ErrorHandler
    
    With thisDb.Relations
        .Append newRelation
    End With
    
ErrorHandler:
    Select Case Err.Number
    Case 0
    Case 3012
        Debug.Print "Relationship already exists: """ & newRelation.Name & """ "
    Case 3284
        Debug.Print "Index already exists for: """ & newRelation.Name & """ "
    Case Else
        Debug.Print "Failed to add: """ & newRelation.Name & """ " & Err.Number & " " & Err.Description
    End Select
End Sub


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
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Collection
    Set IDbComponent_GetFileList = GetFilePathsInFolder(IDbComponent_BaseFolder & "*.json")
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_ClearOrphanedSourceFiles() As Variant
    ClearFilesByExtension IDbComponent_BaseFolder, "txt"
    ClearOrphanedSourceFiles Me, "json"
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
    ' No modification date available for relations
End Function


'---------------------------------------------------------------------------------------
' Procedure : SourceModified
' Author    : Adam Waller
' Date      : 4/27/2020
' Purpose   : The date/time the source object was modified. In most cases, this would
'           : be the date/time of the source file, but it some cases like SQL objects
'           : the date can be determined through other means, so this function
'           : allows either approach to be taken.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_SourceModified() As Date
    If FSO.FileExists(IDbComponent_SourceFile) Then IDbComponent_SourceModified = FileDateTime(IDbComponent_SourceFile)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "relations"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "relations\"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
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
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetRelationFileName(m_Relation) & ".json"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count() As Long
    IDbComponent_Count = IDbComponent_GetAllFromDB.Count
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