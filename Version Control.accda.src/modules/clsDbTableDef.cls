VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbTableDef"
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

Private m_Table As AccessObject
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

    Dim strFile As String
    Dim dbs As Database
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    Dim dItem As Dictionary
    Dim strHash As String
    Dim strContent As String

    ' Clear any legacy extensions
    ClearFilesByExtension IDbComponent_BaseFolder, "LNKD"
    ClearFilesByExtension IDbComponent_BaseFolder, "bas"
    ClearFilesByExtension IDbComponent_BaseFolder, "tdf"

    ' Get the export file name
    strFile = Nz2(strAlternatePath, IDbComponent_SourceFile)

    ' For internal tables, we can export them as XML.
    If Not IsLinkedTable Then

        ' Save structure in XML format
        VerifyPath strFile
        Perf.OperationStart "App.ExportXML()"
        ' Note that the additional properties are important to accurately reconstruct the table.
        Application.ExportXML acExportTable, m_Table.Name, , strFile, , , , acExportAllTableAndFieldProperties
        Perf.OperationEnd

        ' Rewrite sanitized XML as formatted UTF-8 content
        With New clsSourceParser
            .LoadSourceFile strFile
            DeleteFile strFile
            WriteFile .Sanitize(ectXML), strFile
            strHash = .Hash
        End With

    Else
        ' Linked table - Save as JSON
        Set dbs = CurrentDb
        Set tdf = dbs.TableDefs(m_Table.Name)

        Set dItem = New Dictionary
        With dItem
            .Add "Name", m_Table.Name
            .Add "Connect", SanitizeConnectionString(tdf.Connect)
            .Add "SourceTableName", tdf.SourceTableName
            .Add "Attributes", tdf.Attributes
            ' indexes (Find primary key)
            If IndexAvailable(tdf) Then
                For Each idx In tdf.Indexes
                    If idx.Primary Then
                        ' Add the primary key columns, using brackets just in case the field names have spaces.
                        .Add "PrimaryKey", "[" & MultiReplace(CStr(idx.Fields), "+", vbNullString, ";", "], [") & "]"
                        Exit For
                    End If
                Next idx
            End If
        End With

        ' Write export file.
        strContent = BuildJsonFile(TypeName(Me), dItem, "Linked Table")
        strHash = GetStringHash(strContent, True)
        WriteFile strContent, strFile

    End If

    ' Additional processing when exporting to source folder
    If strAlternatePath = vbNullString Then

        ' Remove any alternate source file in case we have switched formats
        RemoveAlternateFormatSourceFile

        ' Optionally save in SQL format
        If Options.SaveTableSQL Then
            Log.Add "  " & m_Table.Name & " (SQL)", Options.ShowDebug
            SaveTableSqlDef m_Table.Name, IDbComponent_BaseFolder
        End If
    End If

    ' Update index
    VCSIndex.Update Me, IIf(strAlternatePath = vbNullString, eatExport, eatAltExport), strHash

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim blnUseTemp As Boolean
    Dim strTempFile As String
    Dim strName As String

    ' Determine import type from extension
    Select Case LCase$(FSO.GetExtensionName(strFile))

        Case "json"
            If Not ImportLinkedTable(strFile) Then Exit Sub

        Case "xml"
            ' Watch out for errors when importing XML
            LogUnhandledErrors
            On Error Resume Next

            ' The ImportXML function does not properly handle UrlEncoded paths
            blnUseTemp = (InStr(1, strFile, "%") > 0)
            If blnUseTemp Then
                ' Import from (safe) temporary file name.
                strTempFile = GetTempFile
                FSO.CopyFile strFile, strTempFile
                Application.ImportXML strTempFile, acStructureOnly
                DeleteFile strTempFile
            Else
                Application.ImportXML strFile, acStructureOnly
            End If

            ' Importing XML can throw an error if ActiveX security is too restrictive.
            If Catch(31521) Then
                Log.Error eelError, "ActiveX security settings blocked the import of " & _
                "table definition XML file: " & FSO.GetFileName(strFile) & ".", ModuleName(Me) & ".Import"
                Log.Add "Importing XML files requires ActiveX to be enabled."
            Else
                ' Log any other error.
                CatchAny eelError, "Unable to import XML for " & FSO.GetFileName(strFile), ModuleName(Me) & ".Import"
            End If

        Case Else
            ' Unsupported file
            Exit Sub

    End Select

    ' Update index
    strName = GetObjectNameFromFileName(strFile)
    Set m_Table = CurrentData.AllTables(strName)
    VCSIndex.Update Me, eatImport

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

    Dim strName As String
    Dim dRelations As Dictionary

    ' Stage any affected relationships
    strName = GetObjectNameFromFileName(strFile)
    Set dRelations = StageRelations(strName)

    ' Delete any existing table with this name
    Set m_Table = Nothing
    DeleteObjectIfExists acTable, strName

    If FSO.FileExists(strFile) Then
        IDbComponent_Import strFile
        ' Restore any affected relationships
        RestoreRelations dRelations
    Else
        ' If the table is deleted, discard any staged relationships
        ' since we can't restore them when the table does not exist.

        ' Also remove table def from index.
        VCSIndex.Remove Me, strFile
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : StageRelations
' Author    : Adam Waller
' Date      : 6/2/2023
' Purpose   : Return a dictionary of the database relationships involving this table.
'           : Each relation is represented by a string value of the JSON source file
'           : content for that relation. (This way we can release the object reference
'           : and delete the actual relation object, but still recreate it later.)
'---------------------------------------------------------------------------------------
'
Private Function StageRelations(strTableName As String) As Dictionary

    Dim dbs As Database
    Dim rel As Relation
    Dim cRel As IDbComponent
    Dim strTempFile As String
    Dim varKey As Variant

    Set StageRelations = New Dictionary

    ' Loop through relations, looking for matching table name
    Set dbs = CurrentDb
    For Each rel In dbs.Relations
        If StrComp(rel.Table, strTableName, vbTextCompare) = 0 Or _
            StrComp(rel.ForeignTable, strTableName, vbTextCompare) = 0 Then
            If strTempFile = vbNullString Then
                ' Create temp file for export
                strTempFile = GetTempFile
                Name strTempFile As strTempFile & ".json"
                strTempFile = strTempFile & ".json"
            End If
            ' Set up DB component class targeting this object
            Set cRel = New clsDbRelation
            Set cRel.DbObject = rel
            cRel.Export strTempFile
            ' Add file content to dictionary
            StageRelations.Add rel.Name, ReadFile(strTempFile)
        End If
    Next rel

    ' Remove the existing relationships affecting this table
    For Each varKey In StageRelations.Keys
        dbs.Relations.Delete varKey
    Next varKey

    ' Remove any temp file
    If FSO.FileExists(strTempFile) Then DeleteFile strTempFile

End Function


'---------------------------------------------------------------------------------------
' Procedure : RestoreRelations
' Author    : Adam Waller
' Date      : 6/2/2023
' Purpose   : Restore any staged relationships.
'---------------------------------------------------------------------------------------
'
Private Function RestoreRelations(dRelations As Dictionary)

    Dim varKey As Variant
    Dim strTempFile As String

    If dRelations.Count > 0 Then
        strTempFile = GetTempFile
        ' Add .json file extension
        Name strTempFile As strTempFile & ".json"
        strTempFile = strTempFile & ".json"
        ' Loop through relations, importing back to database from the temp file.
        For Each varKey In dRelations.Keys
            WriteFile dRelations.Item(varKey), strTempFile
            With New clsDbRelation
                .Parent.Import strTempFile
            End With
        Next varKey
        DeleteFile strTempFile
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveAlternateFormatSourceFile
' Author    : Adam Waller
' Date      : 1/25/2022
' Purpose   : Remove any existing alternate format source file. (When switching a table
'           : from local to linked, or vice versa.)
'---------------------------------------------------------------------------------------
'
Private Sub RemoveAlternateFormatSourceFile()
    Dim strAltFile As String
    strAltFile = IDbComponent_BaseFolder & GetSourceFileName(m_Table.Name, Not IsLinkedTable)
    If FSO.FileExists(strAltFile) Then DeleteFile strAltFile
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveTableSqlDef
' Author    : Adam Waller
' Date      : 1/28/2019
' Purpose   : Save a version of the table formatted as a SQL statement.
'           : (Makes it easier to see table changes in version control systems.)
'---------------------------------------------------------------------------------------
'
Public Sub SaveTableSqlDef(strTable As String, strFolder As String)

    Dim cData As New clsConcat
    Dim cAttr As New clsConcat
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim strFile As String
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef

    Perf.OperationStart "Save Table SQL"

    Set dbs = CurrentDb
    Set tdf = dbs.TableDefs(strTable)

    With cData
        .Add "CREATE TABLE [", DblQ(strTable), "] (", vbCrLf

        ' Loop through fields
        For Each fld In tdf.Fields
            .Add "  [", DblQ(fld.Name), "] "
            If (fld.Attributes And dbAutoIncrField) Then
                .Add "AUTOINCREMENT"
            Else
                .Add GetTypeString(fld.Type), " "
            End If
            Select Case fld.Type
                Case dbText, dbVarBinary
                    .Add "(", fld.Size, ")"
            End Select

            ' Indexes
            For Each idx In tdf.Indexes
                Set cAttr = New clsConcat
                If idx.Fields.Count = 1 And idx.Fields(0).Name = fld.Name Then
                    If idx.Primary Then cAttr.Add " PRIMARY KEY"
                    If idx.Unique Then cAttr.Add " UNIQUE"
                    If idx.Required Then cAttr.Add " NOT NULL"
                    If idx.Foreign Then AddFieldReferences dbs, idx.Fields, strTable, cAttr
                    If Len(cAttr.GetStr) > 0 Then .Add " CONSTRAINT [", DblQ(idx.Name), "]"
                End If
                .Add cAttr.GetStr
            Next
            .Add ",", vbCrLf
        Next fld
        .Remove 3   ' strip off last comma and crlf

        ' Constraints
        If IndexAvailable(tdf) Then
            Set cAttr = New clsConcat
            For Each idx In tdf.Indexes
                If idx.Fields.Count > 1 Then
                    If Len(cAttr.GetStr) = 0 Then cAttr.Add " CONSTRAINT "
                    If idx.Primary Then
                        cAttr.Add "[", DblQ(idx.Name), "] PRIMARY KEY ("
                        For Each fld In idx.Fields
                            cAttr.Add "[", DblQ(fld.Name), "], "
                        Next fld
                        cAttr.Remove 2
                        cAttr.Add ")"
                    End If
                    If Not idx.Foreign Then
                        If Len(cAttr.GetStr) > 0 Then
                            .Add ",", vbCrLf
                            .Add "  ", cAttr.GetStr
                            AddFieldReferences dbs, idx.Fields, strTable, cData
                        End If
                    End If
                End If
            Next idx
        End If
        .Add vbCrLf, ")"

        ' Build file name and create file.
        strFile = strFolder & GetSafeFileName(strTable) & ".sql"
        WriteFile .GetStr, strFile
        Perf.OperationEnd

    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddFieldReferences
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Add references to other fields in table definition.
'---------------------------------------------------------------------------------------
'
Private Sub AddFieldReferences(dbs As Database, fld As Object, strTable As String, cData As clsConcat)

    Dim rel As DAO.Relation
    Dim fld2 As DAO.Field

    For Each rel In dbs.Relations
        If (rel.ForeignTable = strTable) Then
            If FieldsIdentical(fld, rel.Fields) Then

                ' References
                cData.Add " REFERENCES [", DblQ(rel.Table), "] ("
                For Each fld2 In rel.Fields
                    cData.Add "[", DblQ(fld2.Name), "],"
                Next fld2
                ' Remove trailing comma
                If rel.Fields.Count > 0 Then cData.Remove 1
                cData.Add ")"

                ' Attributes for cascade update or delete
                If rel.Attributes And dbRelationUpdateCascade Then cData.Add " ON UPDATE CASCADE "
                If rel.Attributes And dbRelationDeleteCascade Then cData.Add " ON DELETE CASCADE "

                ' Exit now that we have found the matching relationship.
                Exit For

            End If
        End If
    Next rel

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FieldsIdentical
' Author    : Adam Waller
' Date      : 1/21/2019
' Purpose   : Return true if the two collections of fields have the same field names.
'           : (Even if the order of the fields is different.)
'---------------------------------------------------------------------------------------
'
Private Function FieldsIdentical(oFields1 As Object, oFields2 As Object) As Boolean

    Dim fld As Object
    Dim fld2 As Object
    Dim blnMismatch As Boolean
    Dim blnFound As Boolean

    If oFields1.Count <> oFields2.Count Then
        blnMismatch = True
    Else
        ' Set this flag to false after going through each field.
        For Each fld In oFields1
            blnFound = False
            For Each fld2 In oFields2
                If fld.Name = fld2.Name Then
                    blnFound = True
                    Exit For
                End If
            Next fld2
            If Not blnFound Then
                blnMismatch = True
                Exit For
            End If
        Next
    End If

    ' Return result
    FieldsIdentical = Not blnMismatch

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTypeString
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Get the type string used by Access SQL
'---------------------------------------------------------------------------------------
'
Private Function GetTypeString(intType As DAO.DataTypeEnum) As String
    Select Case intType
        Case dbLongBinary:      GetTypeString = "LONGBINARY"
        Case dbBinary:          GetTypeString = "BINARY"
        Case dbBoolean:         GetTypeString = "BIT"
        Case dbAutoIncrField:   GetTypeString = "COUNTER"
        Case dbCurrency:        GetTypeString = "CURRENCY"
        Case dbDate, dbTime:    GetTypeString = "DATETIME"
        Case dbGUID:            GetTypeString = "GUID"
        Case dbMemo:            GetTypeString = "LONGTEXT"
        Case dbDouble:          GetTypeString = "DOUBLE"
        Case dbSingle:          GetTypeString = "SINGLE"
        Case dbByte:            GetTypeString = "UNSIGNED BYTE"
        Case dbInteger:         GetTypeString = "SHORT"
        Case dbLong:            GetTypeString = "LONG"
        Case dbNumeric:         GetTypeString = "NUMERIC"
        Case dbText:            GetTypeString = "VARCHAR"
        Case Else:              GetTypeString = "VARCHAR"
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : IndexAvailable
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return true if the index collection is avilable. Without the error handling
'           : this may throw an error if a linked table is not accessible during export.
'---------------------------------------------------------------------------------------
'
Private Function IndexAvailable(tdf As TableDef) As Boolean

    Dim lngTest As Long

    LogUnhandledErrors
    On Error Resume Next
    lngTest = tdf.Indexes.Count
    If Err Then
        Err.Clear
    Else
        IndexAvailable = True
    End If
    CatchAny eelNoError, vbNullString, , False

End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportLinkedTable
' Author    : Adam Waller
' Date      : 5/6/2020
' Purpose   : Recreate a linked table from the JSON source file.
'           : Returns true if successful.
'---------------------------------------------------------------------------------------
'
Private Function ImportLinkedTable(strFile As String) As Boolean

    Dim dTable As Dictionary
    Dim dItem As Dictionary
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strSql As String
    Dim strConnect As String

    LogUnhandledErrors
    On Error Resume Next

    ' Read json file
    Set dTable = ReadJsonFile(strFile)
    If Not dTable Is Nothing Then

        ' Link the table
        Set dItem = dTable("Items")
        Set dbs = CurrentDb
        Set tdf = dbs.CreateTableDef(dItem("Name"))
        strConnect = GetFullConnect(dItem("Connect"))
        With tdf
            .Connect = strConnect
            .SourceTableName = dItem("SourceTableName")
            .Attributes = SafeAttributes(dItem("Attributes"))
        End With
        dbs.TableDefs.Append tdf
        If Catch(3011, 3024, 3044) Then
            Log.Error eelError, "Could not link table '" & dItem("SourceTableName") & "'", _
            ModuleName(Me) & ".ImportLinkedTable"
            Log.Add "Linked table object not found in " & strFile, False
            Log.Add "Connection String: " & strConnect, False
        ElseIf CatchAny(eelError, vbNullString, ModuleName(Me) & ".ImportLinkedTable") Then
            ' May have encountered other issue like a missing link specification.
        Else
            ' Verify that the connection matches the source file. (Issue #192)
            If tdf.Connect <> strConnect Then
                tdf.Connect = strConnect
                tdf.RefreshLink
            End If
            dbs.TableDefs.Refresh

            ' Set index on linked table.
            If InStr(1, tdf.Connect, ";DATABASE=", vbTextCompare) = 1 Then
                ' Can't create a key on a linked Access database table.
                ' Presumably this would use the Access index instead of needing the pseudo index
            Else
                ' Check for a primary key index (Linked SQL tables may bring over the index, but linked views won't.)
                If dItem.Exists("PrimaryKey") And Not HasUniqueIndex(tdf) Then
                    ' Create a pseudo index on the linked table
                    strSql = "CREATE UNIQUE INDEX __uniqueindex ON [" & DblQ(tdf.Name) & "] (" & dItem("PrimaryKey") & ") WITH PRIMARY"
                    dbs.Execute strSql, dbFailOnError
                    dbs.TableDefs.Refresh
                End If
            End If
            ImportLinkedTable = (Err.Number = 0)
        End If
    End If

    ' Report any unhandled errors
    CatchAny eelError, "Error importing " & strFile, ".ImportLinkedTable"

End Function


'---------------------------------------------------------------------------------------
' Procedure : SafeAttributes
' Author    : Adam Waller
' Date      : 6/29/2020
' Purpose   : Rebuild new attributes flag using attributes that we can actually set.
'---------------------------------------------------------------------------------------
'
Private Function SafeAttributes(lngAttributes As Long) As Long

    Dim colAtts As Collection
    Dim varAtt As Variant
    Dim lngNew As Long

    Set colAtts = New Collection
    With colAtts
        '.Add dbAttachedODBC
        '.Add dbAttachedTable
        .Add dbAttachExclusive
        .Add dbAttachSavePWD
        .Add dbHiddenObject
        .Add dbSystemObject
    End With

    For Each varAtt In colAtts
        ' Use boolean logic to check for bit flag
        If CBool((lngAttributes And varAtt) = varAtt) Then
            ' Add to our rebuilt flag value.
            lngNew = lngNew + varAtt
        End If
    Next varAtt

    ' Return attributes value after rebuilding from scratch.
    SafeAttributes = lngNew

End Function


'---------------------------------------------------------------------------------------
' Procedure : HasUniqueIndex
' Author    : Adam Waller
' Date      : 2/22/2021
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function HasUniqueIndex(tdf As TableDef) As Boolean

    Dim idx As DAO.Index

    If IndexAvailable(tdf) Then
        For Each idx In tdf.Indexes
            If idx.Unique Then
                HasUniqueIndex = True
                Exit For
            End If
        Next idx
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : IDbComponent_MoveSource
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Move the component's source file(s) from one folder to another
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_MoveSource(strFromFolder As String, strToFolder As String)
    MoveFileIfExists strFromFolder & FSO.GetFileName(IDbComponent_SourceFile), strToFolder
    MoveFileIfExists strFromFolder & FSO.GetBaseName(IDbComponent_SourceFile) & ".sql", strToFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary

    Dim tdf As AccessObject
    Dim cTable As IDbComponent
    Dim blnAdd As Boolean

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        Set m_Items(blnModifiedOnly) = New Dictionary
        For Each tdf In CurrentData.AllTables
            If tdf.Name Like "MSys*" Or tdf.Name Like "~*" Then
                ' Skip system and temporary tables
            Else
                Set cTable = New clsDbTableDef
                Set cTable.DbObject = tdf
                blnAdd = True
                If blnModifiedOnly Then blnAdd = cTable.IsModified
                If blnAdd Then m_Items(blnModifiedOnly).Add cTable.SourceFile, cTable
            End If
        Next tdf
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_Items(blnModifiedOnly)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Dictionary
    If m_FileList Is Nothing Then
        Set m_FileList = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.xml")
        MergeDictionary m_FileList, GetFilePathsInFolder(IDbComponent_BaseFolder, "*.json")
    End If
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
    IDbComponent_IsModified = (m_Table.DateModified > VCSIndex.Item(Me).ExportDate)
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
    If m_Table Is Nothing Then Exit Function
    IDbComponent_DateModified = m_Table.DateModified
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "Tables"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "tbldefs" & PathSep
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
    IDbComponent_FileExtensions.Add "xml"
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
    If m_Table Is Nothing Then Exit Property
    IDbComponent_Name = m_Table.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    If m_Table Is Nothing Then Exit Property
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSourceFileName(m_Table.Name, IsLinkedTable)
End Property


'---------------------------------------------------------------------------------------
' Procedure : IsLinkedTable
' Author    : Adam Waller
' Date      : 2/12/2022
' Purpose   : Returns true if this table is a linked table
'---------------------------------------------------------------------------------------
'
Private Property Get IsLinkedTable() As Boolean

    ' Cache the value to save lookups
    Static lngType As Long

    Dim dbs As Database
    Dim rst As Recordset

    ' Nothing to check if we don't have a table name
    If m_Table Is Nothing Then Exit Property

    ' Use cached value if possible
    If lngType > 0 Then
        IsLinkedTable = (lngType <> 1)
        Exit Property
    End If

    Perf.OperationStart "Check for linked table"

    ' This is the simple way to check for a linked table, but it is not very performant
    ' when called hundreds of times while scanning tables for changes.
    'IsLinkedTable = CurrentDb.TableDefs(m_Table.Name).Connect <> vbNullString

    ' A faster way is to read the object type from MSysObjects
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset( _
        "select Type from MSysObjects" & _
        " where name =""" & m_Table.Name & """" & _
        " and Type in(1,4,6)", _
        dbOpenDynaset, dbReadOnly)
    ' Read the table type from the system objects record.
    lngType = Nz(rst!Type, 1)
    rst.Close
    Set rst = Nothing
    Set dbs = Nothing

    ' A value of 1 means a local Access table. Any other value
    ' is almost certainly a linked table of some type. (ODBC, Excel, etc)
    IsLinkedTable = (lngType <> 1)
    Perf.OperationEnd

End Property


'---------------------------------------------------------------------------------------
' Procedure : GetSourceFileName
' Author    : Adam Waller
' Date      : 1/25/2022
' Purpose   : Return a safe file name with the correct file extension.
'---------------------------------------------------------------------------------------
'
Private Function GetSourceFileName(strName As String, blnLinkedTable As Boolean) As String
    If blnLinkedTable Then
        GetSourceFileName = GetSafeFileName(strName) & ".json"
    Else
        GetSourceFileName = GetSafeFileName(strName) & ".xml"
    End If
End Function


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
    If lngCount = 0 Then lngCount = CurrentData.AllTables.Count - 1
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
    IDbComponent_ComponentType = edbTableDef
End Property


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_Table
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Table = RHS
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
