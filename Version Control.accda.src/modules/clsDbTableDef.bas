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

Const ModuleName As String = "clsDbTableDef"

Private m_Table As DAO.TableDef
Private m_AllItems As Collection
Private m_blnModifiedOnly As Boolean
Private m_Dbs As Database

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
    
    Dim strFile As String
    Dim dbs As Database
    Dim tbl As DAO.TableDef
    Dim idx As DAO.Index
    Dim dItem As Dictionary
    
    Set dbs = CurrentDb
    Set tbl = dbs.TableDefs(m_Table.Name)
    strFile = IDbComponent_SourceFile
    
    ' For internal tables, we can export them as XML.
    If tbl.Connect = vbNullString Then
    
        ' Save structure in XML format
        VerifyPath strFile
        Perf.OperationStart "App.ExportXML()"
        ' Note that the additional properties are important to accurately reconstruct the table.
        Application.ExportXML acExportTable, m_Table.Name, , strFile, , , , acExportAllTableAndFieldProperties
        Perf.OperationEnd
        
        ' Rewrite sanitized XML as formatted UTF-8 content
        SanitizeXML strFile
    
    Else
        ' Linked table - Save as JSON
        Set dItem = New Dictionary
        With dItem
            .Add "Name", tbl.Name
            .Add "Connect", SanitizeConnectionString(tbl.Connect)
            .Add "SourceTableName", tbl.SourceTableName
            .Add "Attributes", tbl.Attributes
            ' indexes (Find primary key)
            If IndexAvailable(tbl) Then
                For Each idx In tbl.Indexes
                    If idx.Primary Then
                        ' Add the primary key columns, using brackets just in case the field names have spaces.
                        .Add "PrimaryKey", "[" & MultiReplace(CStr(idx.Fields), "+", vbNullString, ";", "], [") & "]"
                        Exit For
                    End If
                Next idx
            End If
        End With
        
        ' Write export file.
        WriteJsonFile TypeName(Me), dItem, strFile, "Linked Table"
        
    End If
    
    ' Optionally save in SQL format
    If Options.SaveTableSQL Then
        Log.Add "  " & m_Table.Name & " (SQL)", Options.ShowDebug
        SaveTableSqlDef dbs, m_Table.Name, IDbComponent_BaseFolder
    End If

    ' Update index
    VCSIndex.Update Me, eatExport
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveTableSqlDef
' Author    : Adam Waller
' Date      : 1/28/2019
' Purpose   : Save a version of the table formatted as a SQL statement.
'           : (Makes it easier to see table changes in version control systems.)
'---------------------------------------------------------------------------------------
'
Public Sub SaveTableSqlDef(dbs As DAO.Database, strTable As String, strFolder As String)

    Dim cData As New clsConcat
    Dim cAttr As New clsConcat
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim strFile As String
    Dim tdf As DAO.TableDef

    Perf.OperationStart "Save Table SQL"
    Set tdf = dbs.TableDefs(strTable)

    With cData
        .Add "CREATE TABLE [", strTable, "] (", vbCrLf

        ' Loop through fields
        For Each fld In tdf.Fields
            .Add "  [", fld.Name, "] "
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
                    If Len(cAttr.GetStr) > 0 Then .Add " CONSTRAINT [", idx.Name, "]"
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
                        cAttr.Add "[", idx.Name, "] PRIMARY KEY ("
                        For Each fld In idx.Fields
                            cAttr.Add "[", fld.Name, "], "
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
                cData.Add " REFERENCES [", rel.Table, "] ("
                For Each fld2 In rel.Fields
                    cData.Add "[", fld2.Name, "],"
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
    
    If DebugMode(True) Then On Error Resume Next Else On Error Resume Next
    lngTest = tdf.Indexes.Count
    If Err Then
        Err.Clear
    Else
        IndexAvailable = True
    End If
    CatchAny eelNoError, vbNullString, , False
    
End Function


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
        
        Case Else
            ' Unsupported file
            Exit Sub
            
    End Select
    
    ' Update index
    strName = GetObjectNameFromFileName(strFile)
    Set m_Dbs = CurrentDb
    Set m_Table = m_Dbs.TableDefs(strName)
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

End Sub


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
    
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    
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
        If Catch(3011) Then
            Log.Error eelError, "Could not link table '" & dItem("SourceTableName") & "'", _
            ModuleName & ".ImportLinkedTable"
            Log.Add "Linked table object not found in " & strFile, False
            Log.Add "Connection String: " & strConnect, False
        ElseIf CatchAny(eelError, vbNullString, ModuleName & ".ImportLinkedTable") Then
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
                    strSql = "CREATE UNIQUE INDEX __uniqueindex ON [" & tdf.Name & "] (" & dItem("PrimaryKey") & ") WITH PRIMARY"
                    dbs.Execute strSql, dbFailOnError
                    dbs.TableDefs.Refresh
                End If
            End If
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
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Collection
    
    Dim tdf As TableDef
    Dim cTable As IDbComponent
    
    ' Build collection if not already cached
    If m_AllItems Is Nothing Or (blnModifiedOnly <> m_blnModifiedOnly) Then
        Set m_AllItems = New Collection
        m_blnModifiedOnly = blnModifiedOnly
        Set m_Dbs = CurrentDb
        For Each tdf In m_Dbs.TableDefs
            If tdf.Name Like "MSys*" Or tdf.Name Like "~*" Then
                ' Skip system and temporary tables
            Else
                Set cTable = New clsDbTableDef
                Set cTable.DbObject = tdf
                If blnModifiedOnly Then
                    If cTable.IsModified Then m_AllItems.Add cTable, tdf.Name
                Else
                    m_AllItems.Add cTable, tdf.Name
                End If
            End If
        Next tdf
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
    Set IDbComponent_GetFileList = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.xml")
    MergeCollection IDbComponent_GetFileList, GetFilePathsInFolder(IDbComponent_BaseFolder, "*.json")
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_ClearOrphanedSourceFiles()
    ClearFilesByExtension IDbComponent_BaseFolder, "LNKD"
    ClearOrphanedSourceFiles Me, "LNKD", "bas", "sql", "xml", "tdf", "json"
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
    IDbComponent_IsModified = (m_Table.LastUpdated > VCSIndex.GetExportDate(Me))
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
    IDbComponent_DateModified = m_Table.LastUpdated
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
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
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
    If m_Table.Connect = vbNullString Then
        IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Table.Name) & ".xml"
    Else
        ' Linked table
        IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Table.Name) & ".json"
    End If
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
    IDbComponent_ComponentType = edbTableDef
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