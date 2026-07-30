VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbTableData"
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

Public Format As eTableDataExportFormat

Const UNSUPPORTED_DATA_TYPE As String = "UNSUPPORTED DATA TYPE"

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
    Dim strExportedFile As String
    Dim intFormat As eTableDataExportFormat
    Dim strConnect As String

    ' Nothing to export from a category-level instance, such as the one a merge build uses
    ' to re-export after merging a table it ended up skipping.
    If m_Table Is Nothing Then Exit Sub

    ' Use error handling in case we encouter errors during export
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' For linked tables, check if the back-end is already known to be unavailable
    strConnect = SharedDb.TableDefs(m_Table.Name).Connect
    If Len(strConnect) > 0 Then
        If IsBackEndUnavailable(strConnect) Then Exit Sub
    End If

    ' Make sure the table exists, and is accessible
    If Not TableExists(m_Table.Name) Then
        If Len(strConnect) > 0 Then
            If TestBackEndConnection(strConnect) Then
                Log.Error eelError, T("Unable to access table {0} on the back-end database.", _
                    var0:=m_Table.Name), ModuleName(Me) & ".Export"
            Else
                MarkBackEndUnavailable strConnect, m_Table.Name
            End If
        Else
            Log.Error eelError, T("Unable to access table {0}." & _
                " Please ensure that this table exists and is accessible in the current database.", _
                var0:=m_Table.Name), ModuleName(Me) & ".Export"
        End If
        Exit Sub
    End If

    ' Save as selected format, and remove other formats if they exist.
    For intFormat = 1 To eTableDataExportFormat.[_Last] - 1
        ' A caller-supplied path names a single output file, so the cleanup pass for the
        ' other formats would resolve to that same path and delete what we just wrote.
        If Len(strAlternatePath) > 0 Then
            If intFormat <> Me.Format Then GoTo NextFormat
            strFile = strAlternatePath
        Else
            strFile = IDbComponent_BaseFolder & GetSafeFileName(m_Table.Name) & "." & GetExtByFormat(intFormat)
        End If
        If FSO.FileExists(strFile) Then DeleteFile strFile, True
        If intFormat = Me.Format Then
            ' Export the table using this format.
            Select Case intFormat
                Case etdTabDelimited
                    ' Export in tab delimited text file
                    ExportTableDataAsTDF m_Table.Name, strFile
                Case etdXML
                    ' Export data rows as XML (encoding default is UTF-8)
                    VerifyPath strFile
                    ExportTableDataAsXml m_Table.Name, strFile
            End Select
            If Err.Number = 0 Then strExportedFile = strFile
        End If
NextFormat:
    Next intFormat

    ' Record the file we just exported so the next merge sees it as unchanged.
    If Len(strExportedFile) > 0 And Err.Number = 0 Then
        VCSIndex.Update Me, IIf(strAlternatePath = vbNullString, eatExport, eatAltExport), _
            GetFileHash(strExportedFile)
    End If

    ' Trap any errors while exporting table data.
    CatchAny eelError, "Error exporting table data for " & m_Table.Name, ModuleName(Me) & ".Export"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportTableData
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Export the data from the table.
'---------------------------------------------------------------------------------------
'
Private Sub ExportTableDataAsTDF(strTable As String, strFile As String)

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim fld As DAO.Field
    Dim cData As New clsConcat
    Dim intFields As Integer
    Dim intCnt As Integer
    Dim lngProgMax As Long
    Dim lngProgValue As Long

    Perf.OperationStart "Export Table Data as TDF"

    ' Open table in fast read-only view
    Set dbs = SharedDb
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    Set rst = dbs.OpenRecordset(GetTableExportSql(strTable), dbOpenSnapshot, dbReadOnly)
    If Err Then
        Log.Add T("Unable to sort table data for {0} during export. Exporting without ORDER BY.", var0:=strTable), False
        Err.Clear
        Set rst = dbs.OpenRecordset(GetTableExportSql(strTable, False), dbOpenSnapshot, dbReadOnly)
    End If
    CatchAny eelWarning, "Error opening sorted recordset for " & strTable, ModuleName(Me) & ".ExportTableDataAsTDF"
    If rst Is Nothing Then
        Log.Error eelError, T("Unable to export table data for {0}.", var0:=strTable), ModuleName(Me) & ".ExportTableDataAsTDF"
        Exit Sub
    End If
    intFields = rst.Fields.Count

    ' Add header row
    For Each fld In rst.Fields
        cData.Add fld.Name
        intCnt = intCnt + 1
        If intCnt < intFields Then cData.Add vbTab
    Next fld
    cData.Add vbCrLf

    ' Make sure we actually have records
    If Not rst.EOF Then

        ' Populate the recordset so we have an accurate record count
        rst.MoveLast
        rst.MoveFirst

        ' Reuse the console progress bar for the row loop. It is only present when
        ' frmVCSMain is hosting the log, not on web runner or headless test runs.
        If Not Log.ProgressBar Is Nothing Then
            With Log.ProgressBar
                lngProgMax = .Max
                lngProgValue = .Value
                .Max = rst.RecordCount
                .Value = 1
            End With
        End If

        ' Add data rows
        Do While Not rst.EOF
            intCnt = 0
            For Each fld In rst.Fields
                If IsBinaryType(fld) Then
                    ' Binary data not supported
                    cData.Add UNSUPPORTED_DATA_TYPE
                Else
                    ' Format for TDF format without line breaks
                    cData.Add FormatStringForTDF(Nz(fld.Value))
                End If
                intCnt = intCnt + 1
                If intCnt < intFields Then cData.Add vbTab
            Next fld
            cData.Add vbCrLf
            rst.MoveNext
            Log.Increment ' Increment log, in case this takes a while
        Loop

        ' Restore progress bar values for object iteration
        If Not Log.ProgressBar Is Nothing Then
            With Log.ProgressBar
                .Max = lngProgMax
                .Value = lngProgValue
            End With
        End If
    End If

    ' Close database objects
    rst.Close
    Set rst = Nothing

    Perf.OperationEnd

    ' Save output file
    WriteFile cData.GetStr, strFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportTableDataTDF
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Imports the data from a TDF file, loading it into the table
'           : Pass strTargetTable to load the rows somewhere other than the table the
'           : file is named for, such as the staging table used by a merge. The target
'           : is then assumed to be empty, so nothing is cleared first.
'---------------------------------------------------------------------------------------
'
Private Sub ImportTableDataTDF(strFile As String, Optional strTargetTable As String)

    Dim strTable As String
    Dim dCols As Dictionary
    Dim fld As DAO.Field
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim stm As ADODB.Stream
    Dim strLine As String
    Dim varLine As Variant
    Dim varHeader As Variant
    Dim intCol As Integer
    Dim strValue As String

    Perf.OperationStart "Import TDF Data"

    ' Build a dictionary of column names so we can load the data
    ' into the matching columns.
    strTable = Nz2(strTargetTable, GetObjectNameFromFileName(strFile))
    Set dbs = SharedDb
    Set dCols = New Dictionary
    For Each fld In dbs.TableDefs(strTable).Fields
        dCols.Add fld.Name, fld.Name
    Next fld

    ' Clear any existing records before importing this data.
    If Len(strTargetTable) = 0 Then dbs.Execute "delete from [" & strTable & "]", dbFailOnError
    Set rst = dbs.OpenRecordset(strTable)

    ' Read file line by line
    Set stm = New ADODB.Stream
    With stm
        .Charset = "utf-8"
        .Open
        .LoadFromFile strFile
    End With

    ' Loop through lines in file
    Do While Not stm.EOS
        strLine = stm.ReadText(adReadLine)
        ' See if the header has already been parsed.
        If Not IsArray(varHeader) Then
            ' Skip past any UTF-8 BOM header
            If Left$(strLine, 3) = UTF8_BOM Then strLine = Mid$(strLine, 4)
            ' Read header line
            varHeader = Split(strLine, vbTab)
        Else
            ' Data line
            varLine = Split(strLine, vbTab)
            rst.AddNew
                ' Loop through fields
                For intCol = 0 To UBound(varHeader)
                    ' Check to see if field exists in the table
                    If dCols.Exists(varHeader(intCol)) Then
                        ' Check for empty string or null.
                        If varLine(intCol) = vbNullString Then
                            With rst.Fields(varHeader(intCol))
                                If IsNull(.Value) Then
                                    ' Could possibly hit a problem with the storage of
                                    ' zero length strings instead of nulls. Since we can't
                                    ' really differentiate between these in a TDF file,
                                    ' we will do some additional probing...
                                    If .AllowZeroLength And .Required Then
                                        ' Use empty string instead of null value
                                        .Value = vbNullString
                                    End If
                                End If
                            End With
                        Else
                            ' Perform any needed replacements
                            strValue = FormatStringFromTDF(CStr(varLine(intCol)))
                            If strValue <> CStr(varLine(intCol)) Then
                                ' Use replaced string value
                                rst.Fields(varHeader(intCol)).Value = strValue
                            Else
                                If strValue <> UNSUPPORTED_DATA_TYPE Then
                                    ' Use variant value without the string conversion
                                    rst.Fields(varHeader(intCol)).Value = varLine(intCol)
                                End If
                            End If
                        End If
                    End If
                Next intCol
            rst.Update
        End If
        ' Increment log, just in case this takes a while.
        Log.Increment
    Loop
    stm.Close
    Set stm = Nothing
    rst.Close
    Set rst = Nothing

    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FormatStringForTDF
' Author    : Adam Waller
' Date      : 7/16/2021
' Purpose   : Replace line feeds and similar characters with escaped codes for
'           : representation in tab-delimited format.
'           : (Using Chr(26) as interim placeholder) See #251
'---------------------------------------------------------------------------------------
'
Private Function FormatStringForTDF(strValue As String) As String
    FormatStringForTDF = MultiReplace(strValue, _
        "\", Chr$(26), _
        vbCrLf, "\r\n", _
        vbCr, "\r", _
        vbLf, "\n", _
        vbTab, "\t", _
        Chr$(26), "\\")
End Function


'---------------------------------------------------------------------------------------
' Procedure : FormatStringFromTDF
' Author    : Adam Waller
' Date      : 7/16/2021
' Purpose   : Restore original characters from escaped codes.
'---------------------------------------------------------------------------------------
'
Private Function FormatStringFromTDF(strTDFValue) As String
    FormatStringFromTDF = MultiReplace(strTDFValue, _
        "\\", Chr$(26), _
        "\r\n", vbCrLf, _
        "\r", vbCr, _
        "\n", vbLf, _
        "\t", vbTab, _
        Chr$(26), "\")
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExportTableDataAsXml
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Export table data as XML in a deterministic row order.
'           : Application.ExportXML has no ordering parameter and ignores a table's saved
'           : OrderBy property, so the rows come back in engine order. Routing the export
'           : through a temporary ORDER BY query lets the database engine do the sorting,
'           : which avoids reordering the parsed DOM in VBA. Tables that need the embedded
'           : schema, and any failure along the way, fall back to the table export where
'           : clsSourceParser sorts the rows instead.
'---------------------------------------------------------------------------------------
'
Private Sub ExportTableDataAsXml(strTable As String, strFile As String)

    Dim strQuery As String
    Dim strSql As String
    Dim blnExported As Boolean
    Dim blnNeedsSchema As Boolean

    Perf.OperationStart "Check Calculated Fields"
    blnNeedsSchema = TableRequiresXmlSchema(SharedDb.TableDefs(strTable))
    Perf.OperationEnd

    If Not blnNeedsSchema Then
        Perf.OperationStart "Build Table Sort Fields"
        strSql = GetTableExportSql(strTable)
        Perf.OperationEnd
        strQuery = AssignTableDataSortQuery(strSql)
    End If

    If Len(strQuery) > 0 Then
        blnExported = ExportXmlFromQuery(strTable, strQuery, strFile)
        If blnExported Then Exit Sub
    End If

    ExportXmlFromTable strTable, strFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportXmlFromQuery
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Export pre-sorted rows through a temporary query. Returns false when the
'           : caller should fall back to the table export.
'---------------------------------------------------------------------------------------
'
Private Function ExportXmlFromQuery(strTable As String, strQuery As String, strFile As String) As Boolean

    Dim strContent As String
    Dim strFrom As String
    Dim strTo As String

    LogUnhandledErrors
    On Error Resume Next

    Perf.OperationStart "App.ExportXML()"
    If IsLocalTable(strTable) Then
        Application.ExportXML acExportQuery, strQuery, strFile, , , , , acEmbedSchema
    Else
        ' Export data only for linked tables
        Application.ExportXML acExportQuery, strQuery, strFile
    End If
    Perf.OperationEnd

    If Err Then
        Log.Add T("Unable to export sorted table data for {0}. Falling back to an unsorted export.", _
            var0:=strTable), False
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    With New clsSourceParser
        .LoadSourceFile strFile, IDbComponent_ComponentType
        strContent = .Sanitize(ectXML)
    End With
    If Len(strContent) = 0 Then Exit Function

    ' A retained schema would still describe the temporary query rather than the table,
    ' so this path is only safe once sanitization has discarded it. TableRequiresXmlSchema
    ' normally prevents us from getting here, and this guards against it being wrong.
    If InStr(1, strContent, "xsd:schema", vbTextCompare) > 0 Then Exit Function

    ' Row elements carry the temporary query name, so restore the table name to match a
    ' table export. XML escapes "<" inside character data, which means these tags can
    ' only ever appear as markup and never within an exported value.
    strFrom = EscapeXmlName(strQuery)
    strTo = EscapeXmlName(strTable)
    strContent = Replace(strContent, "<" & strFrom & ">", "<" & strTo & ">", , , vbBinaryCompare)
    strContent = Replace(strContent, "</" & strFrom & ">", "</" & strTo & ">", , , vbBinaryCompare)

    ' Never ship a file that still names the temporary query anywhere.
    If InStr(1, strContent, strFrom, vbTextCompare) > 0 Then Exit Function

    WriteFile strContent, strFile
    ExportXmlFromQuery = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExportXmlFromTable
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Export the table directly, letting clsSourceParser reorder the rows.
'---------------------------------------------------------------------------------------
'
Private Sub ExportXmlFromTable(strTable As String, strFile As String)

    Perf.OperationStart "App.ExportXML()"
    If IsLocalTable(strTable) Then
        ' Embedded schema for local tables to support calculated fields
        Application.ExportXML acExportTable, strTable, strFile, , , , , acEmbedSchema
    Else
        ' Export data only for linked tables
        Application.ExportXML acExportTable, strTable, strFile
    End If
    Perf.OperationEnd

    With New clsSourceParser
        ' Loading resets the parser, so the sort fields must be assigned after.
        .LoadSourceFile strFile, IDbComponent_ComponentType
        Set .RowSortFields = GetTableSortFields(SharedDb.TableDefs(strTable))
        WriteFile .Sanitize(ectXML), strFile
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetTableExportSql
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Build SQL to export `tbl_name` sorted by each field from first to last
'---------------------------------------------------------------------------------------
'
Private Function GetTableExportSql(strTable As String, Optional blnIncludeOrderBy As Boolean = True) As String

    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim varKey As Variant
    Dim cText As New clsConcat
    Dim cFieldList As New clsConcat
    Dim cSortList As New clsConcat
    Dim dSortFields As Dictionary
    Dim dbs As Database

    Set dbs = SharedDb
    Set tdf = dbs.TableDefs(strTable)

    ' Build list of fields
    With cFieldList
        For Each fld In tdf.Fields
            .Add "[", fld.Name, "], "
        Next fld
    End With

    ' Remove any trailing commas
    cFieldList.Remove 2

    Set dSortFields = GetTableSortFields(tdf)
    If blnIncludeOrderBy Then
        For Each varKey In dSortFields.Keys
            cSortList.Add "[" & varKey & "], "
        Next varKey
        cSortList.Remove 2
    End If

    ' Build select statement
    With cText
        .Add "SELECT ", cFieldList.GetStr
        .Add " FROM [", strTable, "]"
        If blnIncludeOrderBy And dSortFields.Count > 0 Then
            .Add " ORDER BY "
            .Add cSortList.GetStr
        End If
    End With

    GetTableExportSql = cText.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller, Florian Jenn
' Date      : 4/23/2020, 2020-10-26
' Purpose   : Import the table data from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim blnUseTemp As Boolean
    Dim strTempFile As String
    Dim strTable As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Make sure table exists before importing data to it.
    strTable = GetObjectNameFromFileName(strFile)

    ' Only allow import into local tables (not linked tables)
    If IsLocalTable(strTable) Then

        ' Import from different formats (XML is preferred for data integrity)
        Select Case GetFormatByExt(strFile)
            Case etdXML

                    ' The ImportXML function does not properly handle UrlEncoded paths
                    blnUseTemp = (InStr(1, strFile, "%") > 0)
                    If blnUseTemp Then
                        ' Import from (safe) temporary file name.
                        strTempFile = GetTempFile
                        FSO.CopyFile strFile, strTempFile
                        Application.ImportXML strTempFile, acAppendData
                        DeleteFile strTempFile
                    Else
                        Application.ImportXML strFile, acAppendData
                    End If

                    ' Trap any XML import errors
                    CatchAny eelError, "Error importing XML data from '" & strFile & "'", ModuleName(Me) & ".Import"

            Case etdTabDelimited
                ImportTableDataTDF strFile
        End Select

        ' Record the file we just imported so the next merge sees it as unchanged.
        If Err.Number = 0 And Operation.ErrorLevel <> eelCritical Then
            Set m_Table = CurrentData.AllTables(strTable)
            Me.Format = GetFormatByExt(strFile)
            VCSIndex.Update Me, eatImport, GetFileHash(strFile)
        End If

    Else
        ' Either not a local table, or it doesn't exist
        If TableExists(strTable) Then
            Log.Add T("Skipping table data for linked table '{0}' (data resides in the back-end).", var0:=strTable), False
        Else
            ' Warn user that table does not exist.
            Log.Error eelError, "Table structure not found for '" & strTable & "'.", "clsDbTableData.Import"
            Log.Add "Table definition does not exist for '" & strTable & _
                "'. This must be created before importing table data.", False
            Log.Add "Table data from '" & strTable & "' was not imported.", False
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Merge
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Merge the source file into the existing database, reconciling the rows in
'           : the table against the rows in the source file.
'           :
'           : Import cannot be reused here. It clears the table and reloads every row,
'           : which fails outright once other tables reference this one, and would
'           : rewrite rows that never changed. Instead the file is loaded into a staging
'           : table and reconciled with three set-based statements, so only the rows that
'           : actually differ are touched and existing key values survive.
'           :
'           : Source is authoritative: rows missing from the file are removed from the
'           : table. This is consistent with the index, which has always treated table
'           : data as something to overwrite rather than something to raise a conflict.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Merge(strFile As String)

    Dim strTable As String
    Dim strStaging As String
    Dim strReason As String
    Dim eStrategy As eTableMergeStrategy
    Dim lngInserted As Long
    Dim lngUpdated As Long
    Dim lngDeleted As Long

    strTable = GetObjectNameFromFileName(strFile)

    ' A missing source file means the table was removed from the list of tables to export
    ' data for, or the file was deleted outright. Discarding the rows is never the right
    ' response to either one, so only the index entry is cleared.
    If Not FSO.FileExists(strFile) Then
        Log.Add T("    No source file for {0}. Leaving existing table data unchanged.", _
            var0:=strTable), Options.ShowDebug
        VCSIndex.Remove Me, strFile
        Exit Sub
    End If

    ' Only local tables hold data we can merge.
    If Not IsLocalTable(strTable) Then
        If TableExists(strTable) Then
            Log.Add T("    Skipping table data for linked table '{0}' (data resides in the back-end).", _
                var0:=strTable), Options.ShowDebug
        Else
            Log.Error eelError, "Table structure not found for '" & strTable & "'.", _
                ModuleName(Me) & ".Merge"
            Log.Add T("Table definition does not exist for '{0}'." & _
                " This must be created before merging table data.", var0:=strTable), False
        End If
        Exit Sub
    End If

    eStrategy = GetMergeStrategy(strTable, strFile, strReason)
    If eStrategy = etmsNone Then
        Log.Error eelWarning, T("Unable to merge table data for {0} ({1})." & _
            " Perform a full build to import this table.", var0:=strTable, var1:=strReason), _
            ModuleName(Me) & ".Merge"
        Exit Sub
    End If

    Perf.OperationStart "Merge Table Data"

    strStaging = CreateTableDataStagingTable(strTable)
    If Len(strStaging) = 0 Then
        Log.Error eelError, T("Unable to create a staging table to merge table data for {0}.", _
            var0:=strTable), ModuleName(Me) & ".Merge"
        Perf.OperationEnd
        Exit Sub
    End If

    If LoadStagingTable(strFile, strTable, strStaging) Then
        If ReconcileTableData(strTable, strStaging, eStrategy, _
            lngInserted, lngUpdated, lngDeleted) Then
            If eStrategy = etmsReload Then
                ' Counting changes is not possible without a key, so report what was done.
                Log.Add "    " & T("{0}: {1} row(s) reloaded (no key to compare on)", _
                    var0:=strTable, var1:=CStr(lngInserted)), True
            Else
                Log.Add "    " & T("{0}: {1} added, {2} changed, {3} removed", _
                    var0:=strTable, var1:=CStr(lngInserted), var2:=CStr(lngUpdated), _
                    var3:=CStr(lngDeleted)), _
                    (lngInserted + lngUpdated + lngDeleted > 0) Or Options.ShowDebug
            End If
            ' Record the file we just merged so the next merge sees it as unchanged.
            Perf.OperationStart "Update Index (Table Data)"
            Set m_Table = CurrentData.AllTables(strTable)
            Me.Format = GetFormatByExt(strFile)
            VCSIndex.Update Me, eatImport, GetFileHash(strFile)
            Perf.OperationEnd
        End If
    End If

    DropTableDataStagingTable strStaging
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetMergeStrategy
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Returns how this table and source file can be merged, setting strReason to a
'           : translated explanation when they cannot. Extends the table-level test in
'           : GetTableMergeStrategy with what the source file itself rules out.
'---------------------------------------------------------------------------------------
'
Private Function GetMergeStrategy(strTable As String, strFile As String, _
    ByRef strReason As String) As eTableMergeStrategy

    ' A format we cannot load would leave the staging table empty, and an empty staging
    ' table reads as "the source has no rows" -- which would empty the live table.
    If GetFormatByExt(strFile) = etdNoData Then
        strReason = T("unrecognized source file format")
        Exit Function
    End If

    GetMergeStrategy = GetTableMergeStrategy(SharedDb.TableDefs(strTable), strReason)
    If GetMergeStrategy = etmsNone Then Exit Function

    ' A retained schema names the table it was exported from, so the row elements cannot
    ' be pointed at the staging table without contradicting it. Tables that need the
    ' schema are already rejected above; this catches a file exported before that was so.
    If GetFormatByExt(strFile) = etdXML Then
        If InStr(1, ReadFile(strFile), "xsd:schema", vbTextCompare) > 0 Then
            strReason = T("source file has an embedded schema")
            GetMergeStrategy = etmsNone
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadStagingTable
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Load the rows from the source file into the (empty) staging table.
'---------------------------------------------------------------------------------------
'
Private Function LoadStagingTable(strFile As String, strTable As String, _
    strStaging As String) As Boolean

    Dim strTempFile As String
    Dim lngSourceRows As Long

    LogUnhandledErrors
    On Error Resume Next

    Select Case GetFormatByExt(strFile)

        Case etdTabDelimited
            ImportTableDataTDF strFile, strStaging

        Case etdXML
            ' ImportXML takes the target table from the row element names rather than from
            ' any parameter, so the rows have to be relabeled to reach the staging table.
            strTempFile = GetTempFile
            Perf.OperationStart "Relabel Staging XML"
            If Not WriteStagingXml(strFile, strTable, strStaging, strTempFile, lngSourceRows) Then
                On Error GoTo 0
                Perf.OperationEnd
                Log.Error eelError, T("Unable to prepare table data from '{0}' for merge.", _
                    var0:=strFile), ModuleName(Me) & ".LoadStagingTable"
                Exit Function
            End If
            Perf.OperationEnd
            ' A file with no rows needs no import: the staging table is already empty,
            ' which is what "the source has no records" means. Skipping the call is not
            ' just an optimization -- ImportXML spends about 95 seconds on a row-less
            ' document, however small, while a 2,000 row file imports in a fraction of a
            ' second. Two empty tables were enough to make a merge look broken.
            If lngSourceRows > 0 Then
                Perf.OperationStart "Import XML to Staging"
                Application.ImportXML strTempFile, acAppendData
                Perf.OperationEnd
            End If
            DeleteFile strTempFile

        Case Else
            ' Guarded by GetMergeStrategy. Never fall through to the reconcile with an
            ' empty staging table, which would read as a source file with no rows.
            On Error GoTo 0
            Exit Function

    End Select

    If Err Then
        On Error GoTo 0
        Log.Error eelError, T("Error reading table data from '{0}' for merge. {1}", _
            var0:=strFile, var1:=Err.Description), ModuleName(Me) & ".LoadStagingTable"
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    ' ImportXML can reject individual rows without raising an error, diverting them to an
    ' import errors table. A short load would look like rows the source no longer has, and
    ' the reconcile would delete them, so the counts have to agree before going any further.
    If GetFormatByExt(strFile) = etdXML Then
        If DCount("*", "[" & strStaging & "]") <> lngSourceRows Then
            Log.Error eelError, T("Only {0} of {1} rows from '{2}' could be read for merge.", _
                var0:=CStr(DCount("*", "[" & strStaging & "]")), var1:=CStr(lngSourceRows), _
                var2:=strFile), ModuleName(Me) & ".LoadStagingTable"
            Exit Function
        End If
    End If

    LoadStagingTable = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteStagingXml
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Write a copy of the exported XML with the row elements renamed to the
'           : staging table, so ImportXML loads the rows there instead of into the live
'           : table. Returns false when the file could not be parsed.
'           :
'           : The rows are rebuilt through the DOM rather than by replacing tag text,
'           : because a table is allowed to have a field with the same name as the table
'           : itself. A blind text replacement would rename that field element too.
'---------------------------------------------------------------------------------------
'
Private Function WriteStagingXml(strFile As String, strTable As String, _
    strStaging As String, strOutputFile As String, ByRef lngRows As Long) As Boolean

    Dim objXml As MSXML2.DOMDocument60
    Dim objRow As MSXML2.IXMLDOMNode
    Dim objField As MSXML2.IXMLDOMNode
    Dim cData As clsConcat
    Dim strRowTag As String

    ' Parsed from the file's content rather than its path, since the DOM treats the path
    ' as a URL and these file names can contain percent-encoded characters.
    Set objXml = New MSXML2.DOMDocument60
    If Not objXml.LoadXML(ReadFile(strFile)) Then Exit Function

    strRowTag = EscapeXmlName(strStaging)
    Set cData = New clsConcat
    cData.Add "<?xml version=""1.0"" encoding=""UTF-8""?>", vbCrLf
    ' The namespace is declared even though data-only rows do not normally use it, so a
    ' field element that does carry an od: attribute still resolves.
    cData.Add "<dataroot xmlns:od=""urn:schemas-microsoft-com:officedata"">", vbCrLf

    For Each objRow In objXml.SelectNodes("/dataroot/" & EscapeXmlName(strTable))
        cData.Add "<", strRowTag, ">"
        For Each objField In objRow.ChildNodes
            cData.Add objField.XML
        Next objField
        cData.Add "</", strRowTag, ">", vbCrLf
        lngRows = lngRows + 1
    Next objRow

    cData.Add "</dataroot>"

    WriteFile cData.GetStr, strOutputFile
    WriteStagingXml = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReconcileTableData
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Bring the live table in line with the staging table, returning the number of
'           : rows affected. With a merge key that is three set-based statements; without
'           : one it is a wholesale replace, and lngInserted reports the rows reloaded.
'           :
'           : Order matters. Inserting before deleting means a row that a child record
'           : depends on is present before anything tries to remove its predecessor, and
'           : deleting last means a delete blocked by referential integrity is reported
'           : against a table that is otherwise already up to date.
'           :
'           : The whole reconcile runs in one transaction. A delete blocked by
'           : referential integrity is the expected failure, and rolling back leaves the
'           : table exactly as it was rather than half merged. Errors are always trapped
'           : here, including in debug mode, since an untrapped error would leave the
'           : transaction open.
'---------------------------------------------------------------------------------------
'
Private Function ReconcileTableData(strTable As String, strStaging As String, _
    eStrategy As eTableMergeStrategy, ByRef lngInserted As Long, _
    ByRef lngUpdated As Long, ByRef lngDeleted As Long) As Boolean

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim dKey As Dictionary
    Dim varKey As Variant
    Dim cTargetCols As clsConcat
    Dim cSourceCols As clsConcat
    Dim cJoin As clsConcat
    Dim cSet As clsConcat
    Dim cDiff As clsConcat
    Dim strT As String
    Dim strS As String
    Dim strFirstKey As String
    Dim blnInTrans As Boolean
    Dim blnTiming As Boolean

    Set dbs = SharedDb
    Set tdf = dbs.TableDefs(strTable)
    Set dKey = GetTableMergeKey(tdf)
    strT = "[" & strTable & "]"
    strS = "[" & strStaging & "]"

    Set cTargetCols = New clsConcat
    Set cSourceCols = New clsConcat
    Set cSet = New clsConcat
    Set cDiff = New clsConcat
    For Each fld In tdf.Fields
        cTargetCols.Add "[", fld.Name, "], "
        cSourceCols.Add strS, ".[", fld.Name, "], "
        If Not dKey.Exists(fld.Name) Then
            cSet.Add strT, ".[", fld.Name, "] = ", strS, ".[", fld.Name, "], "
            ' Null is not equal to anything, including another Null, so an inequality
            ' test alone would miss every change to or from Null.
            cDiff.Add "((", strT, ".[", fld.Name, "] <> ", strS, ".[", fld.Name, "])"
            cDiff.Add " OR (", strT, ".[", fld.Name, "] IS NULL XOR ", _
                strS, ".[", fld.Name, "] IS NULL)) OR "
        End If
    Next fld
    cTargetCols.Remove 2
    cSourceCols.Remove 2
    If cSet.Length > 0 Then cSet.Remove 2
    If cDiff.Length > 0 Then cDiff.Remove 4

    Set cJoin = New clsConcat
    For Each varKey In dKey.Keys
        If Len(strFirstKey) = 0 Then strFirstKey = CStr(varKey)
        cJoin.Add "(", strS, ".[", CStr(varKey), "] = ", strT, ".[", CStr(varKey), "]) AND "
    Next varKey
    If cJoin.Length > 0 Then cJoin.Remove 5

    LogUnhandledErrors
    On Error GoTo Failed

    DBEngine.Workspaces(0).BeginTrans
    blnInTrans = True

    If eStrategy = etmsReload Then
        ' No key means no way to pair a source row with a table row, so the rows are
        ' replaced outright. Both statements are inside the transaction, so a failure
        ' leaves the original rows in place rather than an emptied table.
        blnTiming = True
        Perf.OperationStart "Reload: Delete All"
        dbs.Execute "DELETE FROM " & strT, dbFailOnError
        lngDeleted = dbs.RecordsAffected
        Perf.OperationEnd
        Perf.OperationStart "Reload: Insert All"
        dbs.Execute "INSERT INTO " & strT & " (" & cTargetCols.GetStr & ") SELECT " & _
            cSourceCols.GetStr & " FROM " & strS, dbFailOnError
        lngInserted = dbs.RecordsAffected
        Perf.OperationEnd
        blnTiming = False
        DBEngine.Workspaces(0).CommitTrans
        blnInTrans = False
        On Error GoTo 0
        ReconcileTableData = True
        Exit Function
    End If

    blnTiming = True
    Perf.OperationStart "Reconcile: Insert"
    dbs.Execute "INSERT INTO " & strT & " (" & cTargetCols.GetStr & ") SELECT " & _
        cSourceCols.GetStr & " FROM " & strS & " LEFT JOIN " & strT & " ON " & cJoin.GetStr & _
        " WHERE " & strT & ".[" & strFirstKey & "] IS NULL", dbFailOnError
    lngInserted = dbs.RecordsAffected
    Perf.OperationEnd

    ' Nothing to compare or assign when the key covers every column.
    If cSet.Length > 0 Then
        Perf.OperationStart "Reconcile: Update"
        dbs.Execute "UPDATE " & strT & " INNER JOIN " & strS & " ON " & cJoin.GetStr & _
            " SET " & cSet.GetStr & " WHERE " & cDiff.GetStr, dbFailOnError
        lngUpdated = dbs.RecordsAffected
        Perf.OperationEnd
    End If

    Perf.OperationStart "Reconcile: Delete"
    dbs.Execute "DELETE FROM " & strT & " WHERE NOT EXISTS (SELECT 1 FROM " & strS & _
        " WHERE " & cJoin.GetStr & ")", dbFailOnError
    lngDeleted = dbs.RecordsAffected
    Perf.OperationEnd
    blnTiming = False

    DBEngine.Workspaces(0).CommitTrans
    blnInTrans = False
    On Error GoTo 0

    ReconcileTableData = True
    Exit Function

Failed:
    ' Close the statement's timer, not the caller's, so the failure is attributed to the
    ' phase that raised it.
    If blnTiming Then Perf.OperationEnd
    Log.Error eelError, T("Unable to merge table data for {0}. No changes were made" & _
        " to this table. {1}", var0:=strTable, var1:=Err.Description), _
        ModuleName(Me) & ".ReconcileTableData"
    Err.Clear
    ' Clearing the error object is not enough while a handler is active; the handler state
    ' has to be reset before Resume Next protects the rollback.
    On Error GoTo -1
    On Error Resume Next
    If blnInTrans Then DBEngine.Workspaces(0).Rollback
    If Err Then Err.Clear
    On Error GoTo 0

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
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary

    Dim tbl As AccessObject
    Dim cTable As clsDbTableData
    Dim blnNeedAll As Boolean

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        blnNeedAll = (m_Items(False) Is Nothing)
        If blnNeedAll Then Set m_Items(False) = New Dictionary
        If blnModifiedOnly Then Set m_Items(True) = New Dictionary

        ' No need to go any further if we don't have any saved tables defined
        If Options.TablesToExportData.Count > 0 Then

            ' We have at least one table defined. Loop through the tables looking
            ' for a matching name.
            With Options
                For Each tbl In CurrentData.AllTables
                    If .TablesToExportData.Exists(tbl.Name) Then
                        Set cTable = New clsDbTableData
                        cTable.Format = .GetTableExportFormat(CStr(.TablesToExportData(tbl.Name)("Format")))
                        Set cTable.Parent.DbObject = tbl
                        If blnNeedAll Then m_Items(False).Add cTable.Parent.SourceFile, cTable
                        If blnModifiedOnly Then
                            If cTable.IDbComponent_IsModified Then m_Items(True).Add cTable.Parent.SourceFile, cTable
                        End If
                    End If
                Next tbl
            End With
        End If
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_Items(blnModifiedOnly)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type. (Could be
'           : a couple different file extensions involved.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Dictionary
    If m_FileList Is Nothing Then
        Set m_FileList = GetFilePathsInFolder(IDbComponent_BaseFolder, _
            "*." & GetExtByFormat(etdTabDelimited), "*." & GetExtByFormat(etdXML))
    End If
    Set IDbComponent_GetFileList = m_FileList
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsBinaryType
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Returns true if the field type uses binary content. (Export/import on
'           : binary data through XML or text is not well supported.)
'---------------------------------------------------------------------------------------
'
Private Function IsBinaryType(fld As DAO.Field) As Boolean
    IsBinaryType = IsBinaryTableFieldType(fld.Type)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetExtByFormat
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the expected file extension by format.
'---------------------------------------------------------------------------------------
'
Private Function GetExtByFormat(intFormat As eTableDataExportFormat) As String
    Select Case intFormat
        Case etdTabDelimited:   GetExtByFormat = "txt"
        Case etdXML:            GetExtByFormat = "xml"
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFormatByExt
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Look up the format from the extension name
'---------------------------------------------------------------------------------------
'
Private Function GetFormatByExt(strFile As String) As eTableDataExportFormat
    Select Case FSO.GetExtensionName(strFile)
        Case "txt": GetFormatByExt = etdTabDelimited
        Case "xml": GetFormatByExt = etdXML
    End Select
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
    ' It's not really worth it to try to determine if the table data has been changed.
    ' We would have to fully export anyway just to check, so let's always export for
    ' table data. (This is handled different from just about every other object type.)
    IDbComponent_IsModified = True
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
    ' We cannot determine when *records* were modified in a table.
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
    IDbComponent_Category = "Table Data"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "tables" & PathSep
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
    IDbComponent_FileExtensions.Add "txt"
    IDbComponent_FileExtensions.Add "tdf"
    IDbComponent_FileExtensions.Add "xml"
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
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Table.Name) & "." & GetExtByFormat(Me.Format)
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
    ' In most cases very few tables will be involved here
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
    IDbComponent_ComponentType = edbTableData
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
