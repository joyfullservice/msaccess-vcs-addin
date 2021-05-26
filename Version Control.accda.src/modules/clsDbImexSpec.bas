VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbImexSpec"
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


Public Name As String
Public ID As String

Private m_AllItems As Collection

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
    
    Dim dSpec As Dictionary
    Dim dCol As Dictionary
    Dim dCols As Dictionary
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSql As String
    Dim fld As DAO.Field
    
    Set dSpec = New Dictionary
    Set dCols = New Dictionary
    Set dbs = CurrentDb

    ' Build header info first
    strSql = "SELECT * FROM MSysIMEXSpecs WHERE SpecID=" & Me.ID
    Set rst = dbs.OpenRecordset(strSql, dbOpenSnapshot, dbReadOnly)
    With rst
        For Each fld In .Fields
            If fld.Name <> "SpecID" Then
                ' Add all columns except the primary key.
                dSpec.Add CStr(fld.Name), fld.Value
            End If
        Next fld
        .Close
    End With
    
    ' Build list of columns
    strSql = "SELECT * FROM MSysIMEXColumns WHERE SpecID=" & Me.ID
    Set rst = dbs.OpenRecordset(strSql, dbOpenSnapshot, dbReadOnly)
    With rst
        Do While Not .EOF
            Set dCol = New Dictionary
            For Each fld In .Fields
                Select Case fld.Name
                    Case "SpecID", "FieldName"
                    Case Else
                        dCol.Add CStr(fld.Name), fld.Value
                End Select
            Next fld
            ' Add column to columns dictionary
            dCols.Add Nz(!FieldName), dCol
            .MoveNext
        Loop
        .Close
    End With
    
    ' Add columns to spec
    dSpec.Add "Columns", dCols
    
    ' Write as Json format.
    WriteJsonFile TypeName(Me), dSpec, IDbComponent_SourceFile, "Import/Export Specification from MSysIMEXSpecs"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim dFile As Dictionary
    Dim dSpec As Dictionary
    Dim dCol As Dictionary
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim fld As DAO.Field
    Dim lngID As Long
    Dim varKey As Variant

    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub

    ' Read data from JSON file
    Set dFile = ReadJsonFile(strFile)
    If Not dFile Is Nothing Then
    
        ' Create IMEX tables if needed
        VerifyImexTables
        Set dSpec = dFile("Items")
        Set dbs = CurrentDb
        
        ' Add header record
        Set rst = dbs.OpenRecordset("MSysIMEXSpecs")
        With rst
            .AddNew
                For Each fld In .Fields
                    If dSpec.Exists(fld.Name) Then
                        fld.Value = dSpec(fld.Name)
                    End If
                Next fld
            .Update
            ' Save ID from header so we can use it for columns
            .Bookmark = .LastModified
            lngID = !SpecID
        End With
    
        ' Add columns records
        Set rst = dbs.OpenRecordset("MSysIMEXColumns")
        With rst
            For Each varKey In dSpec("Columns").Keys
                Set dCol = dSpec("Columns")(varKey)
                .AddNew
                    !SpecID = lngID
                    !FieldName = CStr(varKey)
                    For Each fld In .Fields
                        If dCol.Exists(fld.Name) Then
                            fld.Value = dCol(fld.Name)
                        End If
                    Next fld
                .Update
            Next varKey
        End With
    
    End If
    
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
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Collection
    
    Dim cSpec As clsDbImexSpec
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim strName As String
    
    ' Build collection if not already cached
    If m_AllItems Is Nothing Then
        
        ' Set up new collection
        Set m_AllItems = New Collection
        
        ' This table may not (yet) exist.
        If TableExists("MSysIMEXSpecs") Then
            ' Look up specs from table
            Set dbs = CurrentDb
            Set rst = dbs.OpenRecordset("MSysIMEXSpecs", dbOpenSnapshot, dbReadOnly)
            With rst
                Do While Not .EOF
                    ' Keep in mind that the spec name may be blank
                    strName = Nz(!SpecName)
                    If strName = vbNullString Then strName = "Spec " & Nz(!SpecID, 0)
                    ' Add spec name
                    Set cSpec = New clsDbImexSpec
                    cSpec.Name = strName
                    cSpec.ID = Nz(!SpecID, 0)
                    m_AllItems.Add cSpec, cSpec.Name
                    .MoveNext
                Loop
                .Close
            End With
        End If
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_AllItems
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : VerifyImexTables
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : If we have not done any import/export in this database, we may need
'           : to create the system tables. For this we will use an undocumented SysCmd
'           : call to create the tables. I have found this documented a few places
'           : online, and used in the VTools project as well.
'           : https://www.everythingaccess.com/tutorials.asp?ID=Undocumented-SysCmd-Functions
'---------------------------------------------------------------------------------------
'
Private Sub VerifyImexTables()
    ' Check to see if the tables exists in the current database
    If (Not TableExists("MSysIMEXSpecs")) Or (Not TableExists("MSysIMEXColumns")) Then
        ' Use an undocumented SysCmd function to create the tables.
        SysCmd 555
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList(Optional blnModifiedOnly As Boolean = False) As Collection
    Set IDbComponent_GetFileList = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.json")
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_ClearOrphanedSourceFiles()
    ClearOrphanedSourceFiles Me, "json"
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
    ' No modified date on a spec
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
    IDbComponent_Category = "IMEX Specs"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "imexspecs" & PathSep
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    IDbComponent_Name = Me.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(Me.Name) & ".json"
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
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbImexSpec
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
