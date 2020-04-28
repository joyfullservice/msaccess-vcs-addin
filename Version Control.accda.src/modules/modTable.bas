Option Explicit
Option Compare Database
Option Private Module


'' Structure to keep track of "on Update" and "on Delete" clauses
'' Access does not in all cases execute such queries
'Private Type structEnforce
'    foreignTable As String
'    foreignFields() As String
'    table As String
'    refFields() As String
'    isUpdate As Boolean
'End Type
'
'' keeping "on Update" relations to be complemented after table creation
'Private K() As structEnforce


'---------------------------------------------------------------------------------------
' Procedure : ExportLinkedTable
' Author    : Adam Waller
' Date      : 1/21/2019
' Purpose   : Export the definition of a linked table
'---------------------------------------------------------------------------------------
'
Public Sub ExportLinkedTable(strTable As String, strFolder As String, cOptions As clsOptions)
    
    Dim blnSkip As Boolean
    Dim strPath As String
    Dim cData As New clsConcat
    Dim dbs As Database
    Dim tdf As DAO.TableDef
    Dim varText As Variant
    Dim idx As DAO.Index
    
    ' Build path
    strPath = strFolder & GetSafeFileName(strTable) & ".LNKD"
    
    ' Check for fast save
    'If cOptions.UseFastSave Then blnSkip = Not (HasMoreRecentChanges(CurrentData.AllTables(strTable), strPath))
    
    ' Export linked table definition
    If blnSkip Then
        If cOptions.ShowDebug Then Debug.Print "  (Skipping '" & strTable & "')"
    Else
        ' Make sure folder exists
        MkDirIfNotExist strFolder
        
        ' Build data string
        With cData
            Set dbs = CurrentDb
            Set tdf = dbs.TableDefs(strTable)
            .Add strTable
            .Add vbCrLf
        
            ' Check for linked databases in the same folder.
            If InStr(1, tdf.connect, "DATABASE=" & CurrentProject.Path) Then
                ' Use relative path for databases in same folder.
                varText = Split(tdf.connect, CurrentProject.Path)
                .Add CStr(varText(0))
                .Add "."
                .Add CStr(varText(1))
            Else
                ' Other folder or link type
                .Add tdf.connect
            End If
            
            ' Source table
            .Add vbCrLf
            .Add tdf.SourceTableName
            .Add vbCrLf
            
            ' Make sure we can access the index.
            ' (Will throw an error if the linked table is not accessible.)
            On Error Resume Next
            varText = tdf.Indexes.Count
            If Err.Number = 3011 Then
                ' File may be inaccessible
                Err.Clear
            ElseIf Err.Number = 3625 Then
                ' Invalid file specification
                Err.Clear
            ElseIf Err.Number > 0 Then
                MsgBox "Error reading linked table indexes on " & tdf.Name & vbCrLf & "Error " & Err.Number & ": " & Err.Description, vbExclamation
                Err.Clear
            Else
                On Error GoTo 0
                ' Indexes
                For Each idx In tdf.Indexes
                    If idx.Primary Then
                        .Add Mid(idx.Fields, 2)
                        .Add vbCrLf
                    End If
                Next idx
            End If
            On Error GoTo 0
            
            ' Write to file
            WriteFile cData.GetStr, strPath
        End With
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportTableDef
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Save a Table Definition as SQL statement.
'           : An alternative would be to export the def as XML, but I feel that this
'           : format is more readable from a code review standpoint.
'---------------------------------------------------------------------------------------
'
Public Sub ExportTableDef(strTable As String, strFolder As String, cOptions As clsOptions)

    Dim strFile As String
    Dim cData As New clsConcat
    Dim blnSkip As Boolean
    Dim dbs As Database
    Dim tdf As TableDef

    Set dbs = CurrentDb
    Set tdf = dbs.TableDefs(strTable)

    ' Build file name
    strFile = strFolder & GetSafeFileName(strTable) & ".xml"

    ' Check for fast save
    'If cOptions.UseFastSave Then blnSkip = Not (HasMoreRecentChanges(CurrentData.AllTables(strTable), strFile))

    ' Export table definition
    If blnSkip Then
        Log.Add "  (Skipping '" & strTable & "')", cOptions.ShowDebug
    Else
        If cOptions.SaveTableSQL Then
            ' Option for SQL output for accdb tables
            Log.Add "  " & strTable & " (with SQL)", cOptions.ShowDebug
            SaveTableSqlDef dbs, strTable, strFolder, cOptions
        Else
            Log.Add "  " & strTable, cOptions.ShowDebug
        End If
        
        ' Tables are export as XML files
        Application.ExportXML acExportTable, strTable, , strFile

        'exort Data Macros
        ExportDataMacros strTable, strFolder, cOptions
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveTableSqlDef
' Author    : Adam Waller
' Date      : 1/28/2019
' Purpose   : Save a version of the table formatted as a SQL statement.
'           : (Makes it easier to see table changes in version control systems.)
'---------------------------------------------------------------------------------------
'
Public Sub SaveTableSqlDef(dbs As DAO.Database, strTable As String, strFolder As String, cOptions As clsOptions)

    Dim cData As New clsConcat
    Dim cAttr As New clsConcat
    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim strFile As String
    Dim tdf As DAO.TableDef
    
    Set tdf = dbs.TableDefs(strTable)
    
    With cData
        .Add "CREATE TABLE ["
        .Add strTable
        .Add "] ("
        .Add vbCrLf
        
        ' Loop through fields
        For Each fld In tdf.Fields
            .Add "  ["
            .Add fld.Name
            .Add "] "
            If (fld.Attributes And dbAutoIncrField) Then
                .Add "AUTOINCREMENT"
            Else
                .Add GetTypeString(fld.Type)
                .Add " "
            End If
            Select Case fld.Type
                Case dbText, dbVarBinary
                    .Add "("
                    .Add fld.Size
                    .Add ")"
            End Select
            
            ' Indexes
            For Each idx In tdf.Indexes
                Set cAttr = New clsConcat
                If idx.Fields.Count = 1 And idx.Fields(0).Name = fld.Name Then
                    If idx.Primary Then cAttr.Add " PRIMARY KEY"
                    If idx.Unique Then cAttr.Add " UNIQUE"
                    If idx.Required Then cAttr.Add " NOT NULL"
                    If idx.Foreign Then AddFieldReferences dbs, idx.Fields, strTable, cAttr
                    If Len(cAttr.GetStr) > 0 Then
                        .Add " CONSTRAINT ["
                        .Add idx.Name
                        .Add "]"
                    End If
                End If
                .Add cAttr.GetStr
            Next
            .Add ","
            .Add vbCrLf
        Next fld
        .Remove 3   ' strip off last comma and crlf

        ' Constraints
        Set cAttr = New clsConcat
        For Each idx In tdf.Indexes
            If idx.Fields.Count > 1 Then
                If Len(cAttr.GetStr) = 0 Then cAttr.Add " CONSTRAINT "
                If idx.Primary Then
                    cAttr.Add "["
                    cAttr.Add idx.Name
                    cAttr.Add "] PRIMARY KEY ("
                    For Each fld In idx.Fields
                        cAttr.Add fld.Name
                        cAttr.Add ", "
                    Next fld
                    cAttr.Remove 2
                    cAttr.Add ")"
                End If
                If Not idx.Foreign Then
                    If Len(cAttr.GetStr) > 0 Then
                        .Add ","
                        .Add vbCrLf
                        .Add "  "
                        .Add cAttr.GetStr
                        AddFieldReferences dbs, idx.Fields, strTable, cData
                    End If
                End If
            End If
        Next
        .Add vbCrLf
        .Add ")"
        
        ' Build file name and create file.
        strFile = strFolder & GetSafeFileName(strTable) & ".sql"
        WriteFile .GetStr, strFile
        
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
                cData.Add " REFERENCES "
                cData.Add rel.Table
                cData.Add " ("
                For Each fld2 In rel.Fields
                    cData.Add fld2.Name
                    cData.Add ","
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
' Procedure : TableExists
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Returns true if the table exists.
'           : Could use the CurrentData.AllTables collection, but this runs
'           : significantly slower.
'---------------------------------------------------------------------------------------
'
Private Function TableExists(strName As String) As Boolean

    Dim dbs As Database
    Dim tdf As TableDef
    
    Set dbs = CurrentDb
    For Each tdf In dbs.TableDefs
        If tdf.Name = strName Then
            TableExists = True
            Exit For
        End If
    Next tdf

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTableExportSql
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Build SQL to export `tbl_name` sorted by each field from first to last
'---------------------------------------------------------------------------------------
'
Private Function GetTableExportSql(strTable As String) As String

    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim intCnt As Integer
    Dim intFields As Integer
    Dim cText As New clsConcat
    Dim cFieldList As New clsConcat
    Dim dbs As Database
    
    Set dbs = CurrentDb
    Set tdf = dbs.TableDefs(strTable)
    intFields = tdf.Fields.Count
    
    ' Build list of fields
    With cFieldList
        For Each fld In tdf.Fields
            .Add "["
            .Add fld.Name
            .Add "]"
            intCnt = intCnt + 1
            If intCnt < intFields Then .Add ", "
        Next fld
    End With
    
    ' Build select statement
    With cText
        .Add "SELECT "
        .Add cFieldList.GetStr
        .Add " FROM ["
        .Add strTable
        .Add "] ORDER BY "
        .Add cFieldList.GetStr
    End With
    
    GetTableExportSql = cText.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExportTableData
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Export the data from the table.
'---------------------------------------------------------------------------------------
'
Public Sub ExportTableData(strTable As String, strFolder As String, cOptions As clsOptions)

    Dim rst As DAO.Recordset
    Dim fld As DAO.Field
    Dim cData As New clsConcat
    Dim intFields As Integer
    Dim intCnt As Integer
    Dim strText As String
    
    ' Make sure table exists
    If Not TableExists(strTable) Then
        Log.Add "Error: Table " & strTable & " missing"
        Exit Sub
    End If
    
    ' Open table in fast read-only view
    Set rst = CurrentDb.OpenRecordset(GetTableExportSql(strTable), dbOpenSnapshot, dbOpenForwardOnly)
    intFields = rst.Fields.Count
    
    ' Add header row
    For Each fld In rst.Fields
        cData.Add fld.Name
        intCnt = intCnt + 1
        If intCnt < intFields Then cData.Add vbTab
    Next fld
    cData.Add vbCrLf

    ' Add data rows
    Do While Not rst.EOF
        intCnt = 0
        For Each fld In rst.Fields
            ' Format for TDF format without line breaks
            strText = MultiReplace(Nz(fld.Value), "\", "\\", vbCrLf, "\n", vbCr, "\n", vbLf, "\n", vbTab, "\t")
            cData.Add strText
            intCnt = intCnt + 1
            If intCnt < intFields Then cData.Add vbTab
        Next fld
        cData.Add vbCrLf
        rst.MoveNext
    Loop
    
    ' Save output file
    MkDirIfNotExist strFolder
    WriteFile cData.GetStr, strFolder & GetSafeFileName(strTable) & ".txt"

End Sub


'===========================================================
'
'       IMPORT FUNCTIONS (Under development)
'
'===========================================================


Public Sub ImportLinkedTable(tblName As String, obj_path As String)
    Dim Db As Database ' DAO.Database
    Dim InFile As Scripting.TextStream
    
    Set Db = CurrentDb
    
    Dim tempFilePath As String
    tempFilePath = GetTempFile()
    
    ConvertUtf8Ucs2 obj_path & tblName & ".LNKD", tempFilePath
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(tempFilePath, ForReading, False, TristateTrue)
    
    On Error GoTo err_notable:
    DoCmd.DeleteObject acTable, tblName
    
    GoTo err_notable_fin:
err_notable:
    Err.Clear
    Resume err_notable_fin:
err_notable_fin:
    On Error GoTo Err_CreateLinkedTable:
    
    Dim td As TableDef
    Set td = Db.CreateTableDef(InFile.ReadLine())
    
    Dim connect As String
    connect = InFile.ReadLine()
    If InStr(1, connect, "DATABASE=.\") Then 'replace relative path with literal path
        connect = Replace(connect, "DATABASE=.\", "DATABASE=" & CurrentProject.Path & "\")
    End If
    td.connect = connect
    
    td.SourceTableName = InFile.ReadLine()
    Db.TableDefs.Append td
    
    GoTo Err_CreateLinkedTable_Fin:
    
Err_CreateLinkedTable:
    MsgBox Err.Description, vbCritical, "ERROR: IMPORT LINKED TABLE"
    Resume Err_CreateLinkedTable_Fin:
Err_CreateLinkedTable_Fin:

    'this will throw errors if a primary key already exists or the table is linked to an access database table
    'will also error out if no pk is present
    On Error GoTo Err_LinkPK_Fin:
    
    Dim Fields As String
    Fields = InFile.ReadLine()
    Dim Field As Variant
    Dim sql As String
    sql = "CREATE INDEX __uniqueindex ON " & td.Name & " ("
    
    For Each Field In Split(Fields, ";+")
        sql = sql & "[" & Field & "]" & ","
    Next
    'remove extraneous comma
    sql = Left(sql, Len(sql) - 1)
    
    sql = sql & ") WITH PRIMARY"
    CurrentDb.Execute sql
    
Err_LinkPK_Fin:
    On Error Resume Next
    InFile.Close
    
End Sub

'
'' Import Table Definition
'Public Sub ImportTableDef(tblName As String, directory As String)
'    Dim filePath As String: filePath = directory & tblName & ".sql"
'    Dim Db As Object ' DAO.Database
'    Dim InFile As Scripting.TextStream
'    Dim buf As String
'    Dim P As Integer
'    Dim p1 As Integer
'    Dim strMsg As String
'    Dim s
'    Dim n As Integer
'    Dim i As Integer
'    Dim j As Integer
'    Dim tempFileName As String: tempFileName = modFileAccess.GetTempFile()
'
'    n = -1
'    modFileAccess.ConvertUtf8Ucs2 filePath, tempFileName
'    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
'    Set InFile = FSO.OpenTextFile(tempFileName, ForReading, False, TristateTrue)
'    Set Db = CurrentDb
'    DoCmd.DeleteObject acTable, tblName
'    buf = InFile.ReadLine()
'    Do Until InFile.AtEndOfStream
'        buf = buf & InFile.ReadLine()
'    Loop
'
'    ' The following block is needed because "on update" actions may cause problems
'    For Each s In Split("UPDATE|DELETE", "|")
'    P = InStr(buf, "ON " & s & " CASCADE")
'    While P > 0
'        n = n + 1
'        ReDim Preserve K(n)
'        K(n).table = tblName
'        K(n).isUpdate = (s = "UPDATE")
'
'        buf = Left(buf, P - 1) & Mid(buf, P + 18)
'        P = InStrRev(buf, "REFERENCES", P)
'        p1 = InStr(P, buf, "(")
'        K(n).foreignFields = Split(SubString(p1, buf, "(", ")"), ",")
'        K(n).foreignTable = Trim(Mid(buf, P + 10, p1 - P - 10))
'        P = InStrRev(buf, "CONSTRAINT", p1)
'        p1 = InStrRev(buf, "FOREIGN KEY", p1)
'        If (p1 > 0) And (P > 0) And (p1 > P) Then
'        ' multifield index
'            K(n).refFields = Split(SubString(p1, buf, "(", ")"), ",")
'        ElseIf p1 = 0 Then
'        ' single field
'        End If
'        P = InStr(P, "ON " & s & " CASCADE", buf)
'    Wend
'    Next
'    On Error Resume Next
'    For i = 0 To n
'        strMsg = K(i).table & " to " & K(i).foreignTable
'        strMsg = strMsg & "(  "
'        For j = 0 To UBound(K(i).refFields)
'            strMsg = strMsg & K(i).refFields(j) & ", "
'        Next j
'        strMsg = Left(strMsg, Len(strMsg) - 2) & ") to ("
'        For j = 0 To UBound(K(i).foreignFields)
'            strMsg = strMsg & K(i).foreignFields(j) & ", "
'        Next j
'        strMsg = Left(strMsg, Len(strMsg) - 2) & ") Check "
'        If K(i).isUpdate Then
'            strMsg = strMsg & " on update cascade " & vbCrLf
'        Else
'            strMsg = strMsg & " on delete cascade " & vbCrLf
'        End If
'    Next
'    On Error GoTo 0
'    Db.Execute buf
'    InFile.Close
'    If Len(strMsg) > 0 Then MsgBox strMsg, vbOKOnly, "Correct manually"
'
'
'End Sub


' Import the lookup table `tblName` from `source\tables`.
Public Sub ImportTableData(tblName As String, obj_path As String)
    Dim Db As Object ' DAO.Database
    Dim rs As Object ' DAO.Recordset
    Dim fieldObj As Object ' DAO.Field
    Dim InFile As Scripting.TextStream
    Dim c As Long, buf As String, Values() As String, Value As Variant
    
    Dim tempFileName As String: tempFileName = modFileAccess.GetTempFile()
    modFileAccess.ConvertUtf8Ucs2 obj_path & tblName & ".txt", tempFileName
    ' open file for reading with Create=False, Unicode=True (USC-2 Little Endian format)
    Set InFile = FSO.OpenTextFile(tempFileName, ForReading, False, TristateTrue)
    Set Db = CurrentDb

    Db.Execute "DELETE FROM [" & tblName & "]"
    Set rs = Db.OpenRecordset(tblName)
    buf = InFile.ReadLine()
    Do Until InFile.AtEndOfStream
        buf = InFile.ReadLine()
        If Len(Trim(buf)) > 0 Then
            Values = Split(buf, vbTab)
            c = 0
            rs.AddNew
            For Each fieldObj In rs.Fields
                DoEvents
                Value = Values(c)
                If Len(Value) = 0 Then
                    Value = Null
                Else
                    Value = Replace(Value, "\t", vbTab)
                    Value = Replace(Value, "\n", vbCrLf)
                    Value = Replace(Value, "\\", "\")
                End If
                rs(fieldObj.Name) = Value
                c = c + 1
            Next
            rs.Update
        End If
    Loop

    rs.Close
    InFile.Close
    FSO.DeleteFile tempFileName
End Sub