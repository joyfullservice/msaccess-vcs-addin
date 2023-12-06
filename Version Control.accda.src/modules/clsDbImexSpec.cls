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
Public ID As Long

Private m_Items(True To False) As Dictionary
Private m_FileList As Dictionary

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
            ' Update class with current ID and name
            Me.ID = lngID
            Me.Name = Nz(!SpecName, "Spec " & lngID)
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

    ' Update index
    VCSIndex.Update Me, eatImport, GetStringHash(GetSource, True)

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

    ' Remove any existing spec by this name
    DeleteSpecByName GetObjectNameFromFileName(strFile)

    If FSO.FileExists(strFile) Then
        IDbComponent_Import strFile
    Else
        ' Remove from index if the source file doesn't exist.
        VCSIndex.Remove Me, strFile
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

    Dim cSpec As clsDbImexSpec
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim strName As String
    Dim blnAdd As Boolean

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then

        ' Set up new collection
        Set m_Items(blnModifiedOnly) = New Dictionary

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
                    blnAdd = True
                    If blnModifiedOnly Then blnAdd = cSpec.IDbComponent_IsModified
                    If blnAdd Then m_Items(blnModifiedOnly).Add cSpec.Parent.SourceFile, cSpec
                    .MoveNext
                Loop
                .Close
            End With
        End If
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_Items(blnModifiedOnly)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSource
' Author    : Adam Waller
' Date      : 2/14/2022
' Purpose   : Return the full content that will be saved to the source file.
'---------------------------------------------------------------------------------------
'
Private Function GetSource() As String
    GetSource = BuildJsonFile(TypeName(Me), GetDictionary, "Import/Export Specification from MSysIMEXSpecs")
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 5/8/2023
' Purpose   : Return a dictionary object with the specification definition.
'---------------------------------------------------------------------------------------
'
Private Function GetDictionary() As Dictionary

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

    ' Make sure we have an object specified
    If Me.ID = 0 Then Exit Function

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

    ' Return dictionary
    Set GetDictionary = dSpec

End Function


'---------------------------------------------------------------------------------------
' Procedure : DeleteSpecByName
' Author    : Adam Waller
' Date      : 5/8/2023
' Purpose   : Delete a saved import specification by name.
'---------------------------------------------------------------------------------------
'
Private Sub DeleteSpecByName(strName As String)

    Dim lngID As Long

    ' Skip lookup if we don't even have the corresponding system table.
    If Not TableExists("MSysIMEXSpecs") Then Exit Sub

    ' Attempt to look up spec by name
    lngID = Nz(DLookup("SpecID", "MSysIMEXSpecs", "SpecName=""" & DblQ(strName) & """"))
    If lngID > 0 Then
        ' Remove related records
        CurrentDb.Execute "delete from MSysIMEXColumns where SpecID=" & lngID, dbFailOnError
        CurrentDb.Execute "delete from MSysIMEXSpecs where SpecID=" & lngID, dbFailOnError
    End If

End Sub


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
Private Function IDbComponent_GetFileList() As Dictionary
    If m_FileList Is Nothing Then Set m_FileList = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.json")
    Set IDbComponent_GetFileList = m_FileList
End Function


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
' Procedure : FileExtensions
' Author    : Adam Waller
' Date      : 12/1/2023
' Purpose   : A collection of the file extensions used in source files for this
'           : component type.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_FileExtensions() As Collection
    Set IDbComponent_FileExtensions = New Collection
    IDbComponent_FileExtensions.Add "json"
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
' Procedure : QuickCount
' Author    : Adam Waller
' Date      : 6/14/2022
' Purpose   : Return a cached, non-iterative approximate count of database objects
'           : for use with progress indicators when scanning for changes. Single file
'           : objects like database properties can simply return 1.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_QuickCount() As Long
    ' No simple way to get count without a database query
    IDbComponent_QuickCount = 1
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
