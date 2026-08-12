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
'@Folder("Components")

Public Name As String
Public Id As Long

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
    Dim lngId As Long
    Dim varKey As Variant

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub

    ' Read data from JSON file
    Set dFile = ReadJsonFile(strFile)
    If dFile Is Nothing Then GoTo CleanUp
    If Not dFile.Exists("Items") Then GoTo CleanUp

    ' Create IMEX tables if needed
    VerifyImexTables
    Set dSpec = dFile("Items")
    If dSpec Is Nothing Then GoTo CleanUp
    Set dbs = SharedDb

    ' Insert the header with an explicit SpecID. Do not write columns unless
    ' this succeeds: a failed header insert would leave the recordset on some
    ' other row, and columns would be attached to that spec.
    lngId = InsertSpecHeader(dSpec)
    If lngId = 0 Then GoTo CleanUp

    Me.Id = lngId
    If dSpec.Exists("SpecName") Then
        Me.Name = GetSpecDisplayName(dSpec("SpecName"), lngId)
    Else
        Me.Name = GetSpecDisplayName(vbNullString, lngId)
    End If

    ' Add columns records
    If dSpec.Exists("Columns") Then
        Set rst = dbs.OpenRecordset("MSysIMEXColumns")
        With rst
            For Each varKey In dSpec("Columns").Keys
                Set dCol = dSpec("Columns")(varKey)
                .AddNew
                    !SpecID = lngId
                    !FieldName = CStr(varKey)
                    For Each fld In .Fields
                        If dCol.Exists(fld.Name) Then
                            fld.Value = dCol(fld.Name)
                        End If
                    Next fld
                .Update
                If CatchAny(eelError, T("Unable to save specification column '{0}' from {1}", _
                    var0:=CStr(varKey), var1:=strFile), ModuleName(Me) & ".Import", True, True) Then GoTo CleanUp
            Next varKey
        End With
    End If

    ' Update index
    VCSIndex.Update Me, eatImport, GetStringHash(GetSource, True)

CleanUp:
    CatchAny eelError, T("Unable to import specification from {0}", var0:=strFile), _
        ModuleName(Me) & ".Import", True, True

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

    Dim lngId As Long

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Replace the existing spec this file represents. Blank SpecName is a
    ' singleton (it is the primary key), so an unnamed source file maps to
    ' the unnamed spec regardless of its current SpecID or file name.
    lngId = ResolveExistingSpecId(strFile)
    If lngId > 0 Then DeleteSpecById lngId

    If FSO.FileExists(strFile) Then
        IDbComponent_Import strFile
    Else
        ' Remove from index if the source file doesn't exist.
        VCSIndex.Remove Me, strFile
    End If

    CatchAny eelError, T("Unable to merge specification from {0}", var0:=strFile), _
        ModuleName(Me) & ".Merge", True, True

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
    Dim blnNeedAll As Boolean

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then

        ' Set up new collection
        blnNeedAll = (m_Items(False) Is Nothing)
        If blnNeedAll Then Set m_Items(False) = New Dictionary
        If blnModifiedOnly Then Set m_Items(True) = New Dictionary

        ' This table may not (yet) exist.
        If TableExists("MSysIMEXSpecs") Then
            ' Look up specs from table
            Set dbs = SharedDb
            Set rst = dbs.OpenRecordset("MSysIMEXSpecs", dbOpenSnapshot, dbReadOnly)
            With rst
                Do While Not .EOF
                    ' Keep in mind that the spec name may be blank
                    strName = GetSpecDisplayName(!SpecName, Nz(!SpecID, 0))
                    ' Add spec name
                    Set cSpec = New clsDbImexSpec
                    cSpec.Name = strName
                    cSpec.Id = Nz(!SpecID, 0)
                    If blnNeedAll Then
                        If m_Items(False).Exists(cSpec.Parent.SourceFile) Then
                            Log.Error eelWarning, T("Duplicate IMEX spec source file '{0}'. Skipping spec '{1}' (ID {2}).", _
                                var0:=FSO.GetFileName(cSpec.Parent.SourceFile), var1:=strName, var2:=cSpec.Id), _
                                ModuleName(Me) & ".GetAllFromDB"
                        Else
                            m_Items(False).Add cSpec.Parent.SourceFile, cSpec
                        End If
                    End If
                    If blnModifiedOnly Then
                        If cSpec.IDbComponent_IsModified Then
                            If Not m_Items(True).Exists(cSpec.Parent.SourceFile) Then
                                m_Items(True).Add cSpec.Parent.SourceFile, cSpec
                            End If
                        End If
                    End If
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

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Set dSpec = New Dictionary
    Set dCols = New Dictionary
    Set dbs = SharedDb

    ' Make sure we have an object specified
    If Me.Id = 0 Then Exit Function

    ' Build header info first
    strSql = "SELECT * FROM MSysIMEXSpecs WHERE SpecID=" & Me.Id
    Set rst = dbs.OpenRecordset(strSql, dbOpenSnapshot, dbReadOnly)
    With rst
        If Not .EOF Then
            ' SpecID is a surrogate, not the primary key (SpecName is). Persist it
            ' under 5.1.0 so unnamed specs keep a stable identity across rebuilds.
            If Options.ExportFormatVersion >= EFV_5_1_0 Then
                dSpec.Add "SpecID", !SpecID
            End If
            For Each fld In .Fields
                If fld.Name <> "SpecID" Then
                    dSpec.Add CStr(fld.Name), fld.Value
                End If
            Next fld
        End If
        .Close
    End With

    ' Build list of columns
    strSql = "SELECT * FROM MSysIMEXColumns WHERE SpecID=" & Me.Id
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

    CatchAny eelError, T("Unable to read specification ID {0}", var0:=Me.Id), _
        ModuleName(Me) & ".GetDictionary", True, True

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSpecDisplayName
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : File-name identity for a spec. Blank SpecName falls back to "Spec {id}".
'           : Used only to pick a file name, never to identify a record on merge.
'---------------------------------------------------------------------------------------
'
Private Function GetSpecDisplayName(varSpecName As Variant, lngId As Long) As String
    Dim strName As String
    strName = Nz(varSpecName)
    If strName = vbNullString Then strName = "Spec " & lngId
    GetSpecDisplayName = strName
End Function


'---------------------------------------------------------------------------------------
' Procedure : ResolveExistingSpecId
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Return the SpecID the source file represents, or 0 if none. Blank
'           : SpecName is the primary key singleton, so an unnamed file maps to the
'           : unnamed spec regardless of autonumber. Named files match SpecName.
'           : Missing files (orphan removal) fall back to the derived file name.
'---------------------------------------------------------------------------------------
'
Private Function ResolveExistingSpecId(strFile As String) As Long

    Dim dFile As Dictionary
    Dim dSpec As Dictionary
    Dim strName As String
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim strFileName As String

    If Not TableExists("MSysIMEXSpecs") Then Exit Function

    If FSO.FileExists(strFile) Then
        Set dFile = ReadJsonFile(strFile)
        If dFile Is Nothing Then Exit Function
        If Not dFile.Exists("Items") Then Exit Function
        Set dSpec = dFile("Items")
        If dSpec Is Nothing Then Exit Function
        If dSpec.Exists("SpecName") Then strName = Nz(dSpec("SpecName"))
        ResolveExistingSpecId = LookupSpecIdByName(strName)
    Else
        ' Orphan removal: the file is gone, so match on the derived file name.
        strFileName = FSO.GetFileName(strFile)
        Set dbs = SharedDb
        Set rst = dbs.OpenRecordset("MSysIMEXSpecs", dbOpenSnapshot, dbReadOnly)
        With rst
            Do While Not .EOF
                If GetSafeFileName(GetSpecDisplayName(!SpecName, Nz(!SpecID, 0))) & ".json" = strFileName Then
                    ResolveExistingSpecId = Nz(!SpecID, 0)
                    Exit Do
                End If
                .MoveNext
            Loop
            .Close
        End With
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : LookupSpecIdByName
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Look up SpecID by SpecName, including the blank-name singleton.
'---------------------------------------------------------------------------------------
'
Private Function LookupSpecIdByName(strName As String) As Long

    Dim rst As DAO.Recordset
    Dim strSql As String

    If strName = vbNullString Then
        strSql = "SELECT SpecID FROM MSysIMEXSpecs WHERE SpecName="""" OR SpecName Is Null"
    Else
        strSql = "SELECT SpecID FROM MSysIMEXSpecs WHERE SpecName=""" & DblQ(strName) & """"
    End If
    Set rst = SharedDb.OpenRecordset(strSql, dbOpenSnapshot, dbReadOnly)
    If Not rst.EOF Then LookupSpecIdByName = Nz(rst!SpecID, 0)
    rst.Close

End Function


'---------------------------------------------------------------------------------------
' Procedure : DeleteSpecById
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Delete a saved import specification and its columns by SpecID.
'---------------------------------------------------------------------------------------
'
Private Sub DeleteSpecById(lngId As Long)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    If lngId <= 0 Then Exit Sub
    If Not TableExists("MSysIMEXSpecs") Then Exit Sub

    SharedDb.Execute "DELETE FROM MSysIMEXColumns WHERE SpecID=" & lngId, dbFailOnError
    SharedDb.Execute "DELETE FROM MSysIMEXSpecs WHERE SpecID=" & lngId, dbFailOnError

    CatchAny eelError, T("Unable to delete specification ID {0}", var0:=lngId), _
        ModuleName(Me) & ".DeleteSpecById", True, True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AllocateSpecId
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Choose an unused SpecID. Prefer the value stored in source when it is
'           : free; otherwise Max(SpecID)+1. Never use the AutoNumber seed: it advances
'           : to last-inserted+1, not max+1, and SpecID has no unique index.
'---------------------------------------------------------------------------------------
'
Private Function AllocateSpecId(dSpec As Dictionary) As Long

    Dim lngId As Long
    Dim rst As DAO.Recordset

    If dSpec.Exists("SpecID") Then
        If IsNumeric(dSpec("SpecID")) Then lngId = CLng(Nz(dSpec("SpecID"), 0))
        If lngId > 0 Then
            Set rst = SharedDb.OpenRecordset( _
                "SELECT COUNT(*) FROM MSysIMEXSpecs WHERE SpecID=" & lngId, dbOpenSnapshot, dbReadOnly)
            If Nz(rst(0), 0) = 0 Then
                rst.Close
                AllocateSpecId = lngId
                Exit Function
            End If
            rst.Close
        End If
    End If

    Set rst = SharedDb.OpenRecordset("SELECT Max(SpecID) FROM MSysIMEXSpecs", dbOpenSnapshot, dbReadOnly)
    AllocateSpecId = Nz(rst(0), 0) + 1
    rst.Close

End Function


'---------------------------------------------------------------------------------------
' Procedure : InsertSpecHeader
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Insert the MSysIMEXSpecs row with an explicit SpecID. Returns the ID, or
'           : 0 if the insert failed. DAO recordsets cannot assign AutoNumber values,
'           : so this uses INSERT INTO, with AddNew only if that statement is rejected.
'---------------------------------------------------------------------------------------
'
Private Function InsertSpecHeader(dSpec As Dictionary) As Long

    Dim lngId As Long
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim rst As DAO.Recordset
    Dim strFields As String
    Dim strValues As String
    Dim strSql As String
    Dim strName As String
    Dim lngCount As Long
    Dim blnInserted As Boolean

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    lngId = AllocateSpecId(dSpec)
    If dSpec.Exists("SpecName") Then strName = Nz(dSpec("SpecName"))
    Set dbs = SharedDb
    Set tdf = dbs.TableDefs("MSysIMEXSpecs")

    strFields = "[SpecID]"
    strValues = CStr(lngId)
    For Each fld In tdf.Fields
        If fld.Name <> "SpecID" Then
            If dSpec.Exists(fld.Name) Then
                strFields = strFields & ",[" & fld.Name & "]"
                strValues = strValues & "," & SqlLiteral(dSpec(fld.Name))
            End If
        End If
    Next fld

    strSql = "INSERT INTO MSysIMEXSpecs (" & strFields & ") VALUES (" & strValues & ")"
    dbs.Execute strSql, dbFailOnError
    If Err.Number = 0 Then
        blnInserted = True
    Else
        ' INSERT INTO may be rejected on some system-table configurations.
        ' Fall back to AddNew, then verify the assigned ID is unique.
        Err.Clear
        Set rst = dbs.OpenRecordset("MSysIMEXSpecs")
        With rst
            .AddNew
                For Each fld In .Fields
                    If fld.Name <> "SpecID" Then
                        If dSpec.Exists(fld.Name) Then fld.Value = dSpec(fld.Name)
                    End If
                Next fld
            .Update
            If Err.Number = 0 Then
                .Bookmark = .LastModified
                lngId = Nz(!SpecID, 0)
                blnInserted = True
            End If
        End With
    End If

    If CatchAny(eelError, T("Unable to save specification '{0}'", var0:=GetSpecDisplayName(strName, lngId)), _
        ModuleName(Me) & ".InsertSpecHeader", True, True) Then Exit Function

    If Not blnInserted Or lngId = 0 Then Exit Function

    Set rst = dbs.OpenRecordset("SELECT COUNT(*) FROM MSysIMEXSpecs WHERE SpecID=" & lngId, dbOpenSnapshot, dbReadOnly)
    lngCount = Nz(rst(0), 0)
    rst.Close
    If lngCount <> 1 Then
        ' SpecID collided. Roll back this header by SpecName (the primary key)
        ' rather than by SpecID, which would also delete the other spec's row.
        dbs.Execute "DELETE FROM MSysIMEXSpecs WHERE SpecName=""" & DblQ(strName) & """", dbFailOnError
        CatchAny eelError, T("Specification ID {0} is already in use; insert was rolled back.", var0:=lngId), _
            ModuleName(Me) & ".InsertSpecHeader", True, True
        Exit Function
    End If

    InsertSpecHeader = lngId

End Function


'---------------------------------------------------------------------------------------
' Procedure : SqlLiteral
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Format a value as an Access SQL literal for INSERT INTO.
'---------------------------------------------------------------------------------------
'
Private Function SqlLiteral(varValue As Variant) As String
    If IsNull(varValue) Then
        SqlLiteral = "NULL"
    Else
        Select Case VarType(varValue)
            Case vbBoolean
                SqlLiteral = IIf(varValue, "True", "False")
            Case vbString
                SqlLiteral = """" & DblQ(CStr(varValue)) & """"
            Case Else
                SqlLiteral = CStr(varValue)
        End Select
    End If
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
Private Property Get IDbComponent_FileExtensions(Optional Scope As eFileExtensionScope = efesIndexed) As Collection
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
