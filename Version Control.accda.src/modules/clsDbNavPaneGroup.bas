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

Private m_AllItems As Collection
Private m_dItems As Dictionary
Private m_Count As Long

' This is used to transfer the details to the class.
Private m_Rst As ADODB.Recordset
' Group properties
Private m_GroupName As String
Private m_GroupFlags As Long
Private m_GroupPosition As Long
' Linked object
Private m_ObjectType As Long
Private m_ObjectName As String
Private m_ObjectFlags As Long
Private m_ObjectIcon As Long
Private m_ObjectPosition As Long


' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export()
    IDbComponent_GetAllFromDB
    WriteJsonFile Me, m_dItems, IDbComponent_SourceFile, "Navigation Pane Custom Groups"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'           : Here we are writing information directly to the system tables since
'           : Microsoft Access does not provide a way to do this programatically.
'           : Helpful links: https://stackoverflow.com/questions/26523619
'           : and https://stackoverflow.com/questions/27366038
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim dFile As Dictionary
    Dim varGroup As Variant
    Dim intGroup As Integer
    Dim dGroup As Dictionary
    Dim lngGroupID As Long
    Dim varObject As Variant
    Dim intObject As Integer
    Dim dObject As Dictionary
    Dim lngObjectID As Long
    Dim lngLinkID As Long
    
    Set dFile = ReadJsonFile(strFile)
    If Not dFile Is Nothing Then
        If dFile("Items").Exists("Groups") Then
            For intGroup = 1 To dFile("Items")("Groups").Count
            'For Each varGroup In dFile("Items")("Groups").Keys
                Set dGroup = dFile("Items")("Groups")(intGroup)
                ' Add additional field values for new record
                dGroup.Add "GroupCategoryID", 3
                dGroup.Add "Object Type Group", -1
                dGroup.Add "ObjectID", 0
                ' Check for existing group with this name. (Such as Unassigned Objects)
                lngGroupID = Nz(DLookup("Id", "MSysNavPaneGroups", "GroupCategoryID=3 AND Name=""" & dGroup("Name") & """"), 0)
                If lngGroupID = 0 Then lngGroupID = LoadRecord("MSysNavPaneGroups", dGroup)
                For intObject = 1 To dGroup("Objects").Count
                'For Each varObject In dGroup("Objects").Keys
                    Set dObject = dGroup("Objects")(intObject)
                    lngObjectID = Nz(DLookup("Id", "MSysObjects", "Name=""" & dObject("Name") & """ AND Type=" & dObject("Type")), 0)
                    If lngObjectID <> 0 Then
                        dObject.Add "ObjectID", lngObjectID
                        dObject.Add "GroupID", lngGroupID
                        ' Change name to the name defined in this group. (Could be different from the object name)
                        dObject("Name") = dObject("NameInGroup")
                        ' Should not already be a link, but just in case...
                        lngLinkID = Nz(DLookup("Id", "MSysNavPaneGroupToObjects", "ObjectID=" & lngObjectID & " AND GroupID = " & lngGroupID), 0)
                        If lngLinkID = 0 Then lngLinkID = LoadRecord("MSysNavPaneGroupToObjects", dObject)
                    End If
                Next intObject
            Next intGroup
        End If
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadRecord
' Author    : Adam Waller
' Date      : 5/12/2020
' Purpose   : Loads a new record into the specified table and returns the ID
'---------------------------------------------------------------------------------------
'
Private Function LoadRecord(strTable As String, dCols As Dictionary) As Long

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim fld As DAO.Field
    
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(strTable)
    With rst
        .AddNew
            For Each fld In .Fields
                ' Load field value in matching column
                If dCols.Exists(fld.Name) Then fld.Value = dCols(fld.Name)
            Next fld
        .Update
        .Bookmark = .LastModified
        ' Return ID from new record.
        LoadRecord = Nz(!ID, 0)
        .Close
    End With
    
    Set rst = Nothing
    Set dbs = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB() As Collection

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSQL As String
    Dim strGroup As String
    Dim colGroups As Collection
    Dim dGroup As Dictionary
    Dim colObjects As Collection
    Dim dObject As Dictionary
    
    ' Build collection if not already cached
    If m_AllItems Is Nothing Then

        Set m_AllItems = New Collection
        Set m_dItems = New Dictionary
        Set colGroups = New Collection
        m_Count = 0
        
        ' Load query SQL from saved query in add-in database
        Set dbs = CodeDb
        strSQL = dbs.QueryDefs("qryNavPaneGroups").SQL
        
        ' Open query in the current db
        Set dbs = CurrentDb
        Set rst = dbs.OpenRecordset(strSQL)
        
        ' Loop through records
        With rst
            Do While Not .EOF
                
                ' Check for change in group name.
                If Nz(!GroupName) <> strGroup Then
                    ' Finish recording any previous group
                    If strGroup <> vbNullString Then
                        dGroup.Add "Objects", colObjects
                        colGroups.Add dGroup
                    End If
                    ' Set up new group
                    Set colObjects = New Collection
                    Set dGroup = New Dictionary
                    strGroup = Nz(!GroupName)
                    dGroup.Add "Name", strGroup
                    dGroup.Add "Flags", Nz(!GroupFlags, 0)
                    dGroup.Add "Position", Nz(!GroupPosition, 0)
                End If
            
                ' Add any item listed in this group
                If Nz(!ObjectName) = vbNullString Then
                    ' Saved group with no items.
                    m_Count = m_Count + 1
                Else
                    Set dObject = New Dictionary
                    dObject.Add "Name", Nz(!ObjectName)
                    dObject.Add "Type", Nz(!ObjectType, 0)
                    dObject.Add "Flags", Nz(!ObjectFlags, 0)
                    dObject.Add "Icon", Nz(!ObjectIcon, 0)
                    dObject.Add "Position", Nz(!ObjectPosition, 0)
                    dObject.Add "NameInGroup", Nz(!NameInGroup)
                    colObjects.Add dObject
                    m_Count = m_Count + 1
                End If
                
                ' Move to next record.
                .MoveNext
            Loop
            .Close
            ' Close out last group and add to items dictionary
            If strGroup <> vbNullString Then
                dGroup.Add "Objects", colObjects
                colGroups.Add dGroup
            End If
            m_dItems.Add "Groups", colGroups
        End With
            
        ' Add reference to this class.
        m_AllItems.Add Me

    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_AllItems

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Collection
    Set IDbComponent_GetFileList = GetFilePathsInFolder(IDbComponent_SourceFile)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_ClearOrphanedSourceFiles()
    If IDbComponent_GetAllFromDB.Count = 0 Then
        If FSO.FileExists(IDbComponent_SourceFile) Then Kill IDbComponent_SourceFile
    End If
End Sub


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
    IDbComponent_Category = "nav pane groups"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    IDbComponent_Name = "Groups"
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & "nav-pane-groups.json"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count() As Long
    IDbComponent_GetAllFromDB
    IDbComponent_Count = m_Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbNavPaneGroup
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
    Set IDbComponent_DbObject = Nothing
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
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
    IDbComponent_SingleFile = True
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