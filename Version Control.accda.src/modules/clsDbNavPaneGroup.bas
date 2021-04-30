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

Private Const ModuleName As String = "clsDbNavPaneGroup"
Private Const FormatVersion As Double = 1.1

Private m_AllItems As Collection
Private m_dItems As Dictionary
Private m_Count As Long

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
    WriteJsonFile TypeName(Me), m_dItems, IDbComponent_SourceFile, "Navigation Pane Custom Groups", FormatVersion
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
'           : https://support.microsoft.com/en-us/office/customize-the-navigation-pane-ccfb0ee7-d72f-4923-b4fb-ed6c15484244
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim dFile As Dictionary
    Dim intGroup As Integer
    Dim intCategory As Integer
    Dim dGroup As Dictionary
    Dim dCategory As Dictionary
    Dim lngGroupID As Long
    Dim lngCategoryID As Long
    Dim intObject As Integer
    Dim dObject As Dictionary
    Dim lngCategory As Long
    Dim lngObjectID As Long
    Dim lngLinkID As Long
    
    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub

    Set dFile = ReadJsonFile(strFile)
    If Not dFile Is Nothing Then
    
        ' Upgrade from any previous formats
        Set m_dItems = dFile
        IDbComponent_Upgrade
        
        ' Remove any existing custom groups (Some may be automatically created with a new database)
        ClearExistingNavGroups
    
        ' Import custom navigation categories/groups
        If m_dItems("Items").Exists("Categories") Then
            
            ' Loop through custom categories
            For intCategory = 1 To m_dItems("Items")("Categories").Count
                Set dCategory = m_dItems("Items")("Categories")(intCategory)
                ' Check for existing category with this name
                lngCategoryID = Nz(DLookup("Id", "MSysNavPaneGroupCategories", "Type=4 and Name=""" & dCategory("Name") & """"), 0)
                If lngCategoryID = 0 Then
                    ' Add additional field values and create record
                    dCategory.Add "Type", 4
                    lngCategoryID = LoadRecord("MSysNavPaneGroupCategories", dCategory)
                End If
                ' Make sure we got a category record
                If lngCategoryID = 0 Then
                    Log.Error eelError, _
                        "Could not create custom category record for " & dCategory("Name") & " in MSysNavPaneGroupCategories.", _
                        ModuleName & ".Import"
                    Exit Sub
                End If
        
                ' Loop through groups in category
                For intGroup = 1 To m_dItems("Items")("Categories")(intCategory)("Groups").Count
                    Set dGroup = m_dItems("Items")("Categories")(intCategory)("Groups")(intGroup)
                    ' Add additional field values for new record
                    dGroup.Add "GroupCategoryID", lngCategoryID
                    dGroup.Add "Object Type Group", -1
                    dGroup.Add "ObjectID", 0
                    ' Check for existing group with this name. (Such as Unassigned Objects)
                    lngGroupID = Nz(DLookup("Id", "MSysNavPaneGroups", "GroupCategoryID=" & lngCategoryID & " AND Name=""" & dGroup("Name") & """"), 0)
                    If lngGroupID = 0 Then lngGroupID = LoadRecord("MSysNavPaneGroups", dGroup)
                    For intObject = 1 To dGroup("Objects").Count
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
            Next intCategory
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearExistingNavGroups
' Author    : Adam Waller
' Date      : 2/22/2021
' Purpose   : Clears existing custom groups/categories (Used before importing)
'---------------------------------------------------------------------------------------
'
Private Sub ClearExistingNavGroups()

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSql As String
    
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    
    ' Get SQL for query of NavPaneGroup objects
    Set dbs = CodeDb
    strSql = dbs.QueryDefs("qryNavPaneGroups").SQL
        
    ' Look up list of custom categories
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(strSql, dbOpenSnapshot)
    
    With rst
        Do While Not .EOF
            ' Remove records from three tables
            If Nz(!LinkID, 0) <> 0 Then dbs.Execute "delete from MSysNavPaneGroupToObjects where id=" & Nz(!LinkID, 0), dbFailOnError
            If Nz(!GroupID, 0) <> 0 Then dbs.Execute "delete from MSysNavPaneGroups where id=" & Nz(!GroupID, 0), dbFailOnError
            If Nz(!CategoryID, 0) <> 0 Then dbs.Execute "delete from MSysNavPaneGroupCategories where id=" & Nz(!CategoryID, 0), dbFailOnError
            .MoveNext
        Loop
        .Close
    End With

    CatchAny eelError, "Error clearing existing navigation pane groups.", ModuleName & ".ClearExisting"
    
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
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Collection

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSql As String
    Dim strCategory As String
    Dim strGroup As String
    Dim colCategories As Collection
    Dim colGroups As Collection
    Dim dCategory As Dictionary
    Dim dGroup As Dictionary
    Dim colObjects As Collection
    Dim dObject As Dictionary
    
    ' Build collection if not already cached
    If m_AllItems Is Nothing Then

        Set m_AllItems = New Collection
        Set m_dItems = New Dictionary
        Set colCategories = New Collection
        Set colGroups = New Collection
        m_Count = 0
        
        ' Load query SQL from saved query in add-in database
        Set dbs = CodeDb
        strSql = dbs.QueryDefs("qryNavPaneGroups").SQL
        
        ' Open query in the current db
        Set dbs = CurrentDb
        Set rst = dbs.OpenRecordset(strSql)
        
        ' Loop through records
        With rst
            Do While Not .EOF
            
                ' Check for change in category name
                If Nz(!CategoryName) <> strCategory Then
                    ' Finish recording any previous category
                    If strCategory <> vbNullString Then
                        dCategory.Add "Groups", colGroups
                        colCategories.Add dCategory
                    End If
                    ' Set up new category
                    'Set colCategories = New Collection
                    Set colGroups = New Collection
                    strCategory = Nz(!CategoryName)
                    Set dCategory = New Dictionary
                    dCategory.Add "Name", strCategory
                    dCategory.Add "Flags", Nz(!CategoryFlags, 0)
                    dCategory.Add "Position", Nz(!CategoryPosition, 0)
                End If
                                   
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
            ' Close out last group and category, and add items
            ' to output dictionary
            If strGroup <> vbNullString Then
                dGroup.Add "Objects", colObjects
                colGroups.Add dGroup
            End If
            If strCategory <> vbNullString Then
                dCategory.Add "Groups", colGroups
                colCategories.Add dCategory
                m_dItems.Add "Categories", colCategories
            End If
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
Private Function IDbComponent_GetFileList(Optional blnModifiedOnly As Boolean = False) As Collection
    Set IDbComponent_GetFileList = New Collection
    If FSO.FileExists(IDbComponent_SourceFile) Then IDbComponent_GetFileList.Add IDbComponent_SourceFile
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
        If FSO.FileExists(IDbComponent_SourceFile) Then DeleteFile IDbComponent_SourceFile, True
    End If
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
    If FSO.FileExists(IDbComponent_SourceFile) Then IDbComponent_SourceModified = GetLastModifiedDate(IDbComponent_SourceFile)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "Nav Pane Groups"
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
Private Property Get IDbComponent_Count(Optional blnModifiedOnly As Boolean = False) As Long
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
    
    Dim dNew As Dictionary
    Dim dNew2 As Dictionary
    Dim colNew As Collection
    Dim dblVersion As Double
    
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    
    ' Get version
    If Not m_dItems Is Nothing Then
        If m_dItems("Info").Exists("Export File Format") Then
            dblVersion = CDbl(m_dItems("Info")("Export File Format"))
        End If
    End If
    
    ' Add Category section (2/22/2021)
    If dblVersion < 1.1 Then
        Set dNew = New Dictionary
        Set colNew = New Collection
        ' Build generic Custom category
        dNew.Add "Name", "Custom"
        dNew.Add "Flags", 0
        dNew.Add "Position", 2
        dNew.Add "Groups", m_dItems("Items")("Groups")
        colNew.Add dNew
        Set dNew2 = New Dictionary
        dNew2.Add "Categories", colNew
        Set m_dItems = New Dictionary
        m_dItems.Add "Items", dNew2
    End If
    
    ' Check for newer export file that this add-in version doesn't support
    If dblVersion > FormatVersion Then
        Log.Error eelError, "Format " & dblVersion & " of " & IDbComponent_SourceFile & _
            " not supported by this version of the add-in. Please update the add-in to import this file.", _
            ModuleName & ".Upgrade"
    End If
    
    ' Report any errors during upgrade process
    CatchAny eelError, "Error upgrading navigation pane groups.", ModuleName & ".Export", True, True

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