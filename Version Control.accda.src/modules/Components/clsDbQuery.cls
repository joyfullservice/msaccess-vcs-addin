VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbQuery"
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
'           :
'           : v5.0+: Exports .sql (source of truth for SQL) + .json (metadata)
'           : by reading MSysQueries and MSysObjects directly, bypassing SaveAsText.
'           : Imports by generating .qdef on-the-fly from .sql + .json and calling
'           : LoadFromText.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Components")

Private m_Query As AccessObject
Private m_Items(True To False) As Dictionary
Private m_FileList As Dictionary

Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the individual database component (table, form, query, etc...)
'           :
'           : v5.0+: Reads MSysQueries rows and MSysObjects blobs directly.
'           :   1. Reconstruct SQL from MSysQueries via clsQueryComposer
'           :   2. Format SQL with clsSqlFormatter
'           :   3. Parse LvProp for properties and column metadata
'           :   4. Parse LvExtra for design layout (if present)
'           :   5. Write .sql and .json files
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export(Optional strAlternatePath As String)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim strBasePath As String
    Dim strSqlFile As String
    Dim strJsonFile As String
    Dim blnAlternate As Boolean

    blnAlternate = (strAlternatePath <> vbNullString)

    If Options.ExportFormatVersion >= EFV_5_0_0 Then
        If blnAlternate Then
            strBasePath = Left$(strAlternatePath, Len(strAlternatePath) - 4)
        Else
            strBasePath = IDbComponent_BaseFolder & GetSafeFileName(m_Query.Name)
        End If
        strSqlFile = strBasePath & ".sql"
        strJsonFile = strBasePath & ".json"

        ExportNewFormat strSqlFile, strJsonFile

        If blnAlternate Then
            VCSIndex.Update Me, eatAltExport, GetFileHash(strSqlFile), _
                strMetaHash:=GetMetadataHash("Tables", m_Query.Name, acQuery)
        Else
            VCSIndex.Update Me, eatExport, GetFileHash(strSqlFile), _
                strMetaHash:=GetMetadataHash("Tables", m_Query.Name, acQuery)
        End If
    Else
        If blnAlternate Then
            ExportLegacy strAlternatePath, blnAlternate:=True
        Else
            ExportLegacy IDbComponent_SourceFile
        End If
    End If

    CatchAny eelError, "Error exporting query '" & m_Query.Name & "'", _
        ModuleName(Me) & ".IDbComponent_Export"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportNewFormat
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Core export logic for the new .sql + .json format.
'---------------------------------------------------------------------------------------
'
Private Sub ExportNewFormat(strSqlFile As String, strJsonFile As String)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim strSql As String
    Dim strQueryName As String
    Dim lngObjectId As Long
    Dim colRows As Collection
    Dim dRow As Dictionary
    Dim bteLvProp() As Byte
    Dim bteLvExtra() As Byte
    Dim blnHasLvExtra As Boolean
    Dim dProps As Dictionary
    Dim dLayout As Dictionary
    Dim dItems As Dictionary
    Dim dFile As Dictionary
    Dim dHeader As Dictionary
    Dim dJsonProps As Dictionary
    Dim dJsonColumns As Dictionary
    Dim varKey As Variant
    Dim dFieldProps As Dictionary
    Dim dPropEntry As Dictionary
    Dim cComposer As clsQueryComposer
    Dim strConnect As String
    Dim strEnvKey As String
    Dim lngQueryType As Long

    strQueryName = m_Query.Name
    Set dbs = SharedDb

    ' Step 1: Get the ObjectId and LvProp/LvExtra from MSysObjects
    Perf.OperationStart "Read MSysObjects"
    Set rst = dbs.OpenRecordset( _
        "SELECT Id, LvProp, LvExtra, Flags FROM MSysObjects " & _
        "WHERE Name=""" & EscapeSqlQuote(strQueryName) & """ AND Type=5", _
        dbOpenSnapshot)
    If rst.EOF Then
        rst.Close
        Set rst = Nothing
        Perf.OperationEnd
        Log.Error eelWarning, "Query '" & strQueryName & "' not found in MSysObjects.", _
            "clsDbQuery.ExportNewFormat"
        GoTo CleanUp
    End If

    lngObjectId = rst!ID
    lngQueryType = Nz(rst!Flags, 0)
    If Not IsNull(rst!LvProp) Then bteLvProp = rst!LvProp
    blnHasLvExtra = Not IsNull(rst!LvExtra)
    If blnHasLvExtra Then bteLvExtra = rst!LvExtra
    rst.Close
    Perf.OperationEnd

    ' Step 2: Read MSysQueries rows for this query
    Perf.OperationStart "Read MSysQueries"
    Set colRows = New Collection
    Set rst = dbs.OpenRecordset( _
        "SELECT Attribute, Flag, Expression, Name1, Name2 " & _
        "FROM MSysQueries WHERE ObjectId=" & lngObjectId & _
        " ORDER BY Attribute, [Order]", _
        dbOpenSnapshot)
    Do While Not rst.EOF
        Set dRow = New Dictionary
        dRow.Add "Attribute", CLng(rst!Attribute)
        dRow.Add "Flag", Nz(rst!Flag, 0)
        dRow.Add "Expression", Nz(rst!Expression, vbNullString)
        dRow.Add "Name1", Nz(rst!Name1, vbNullString)
        dRow.Add "Name2", Nz(rst!Name2, vbNullString)
        colRows.Add dRow
        rst.MoveNext
    Loop
    rst.Close
    Perf.OperationEnd

    ' Step 3: Reconstruct SQL from MSysQueries rows
    Perf.OperationStart "Reconstruct SQL"
    Set cComposer = New clsQueryComposer
    strSql = cComposer.ReconstructSQL(colRows)
    Perf.OperationEnd

    ' Step 4: Get the connection string for pass-through queries
    strConnect = vbNullString
    Dim dR As Dictionary
    For Each dR In colRows
        If CLng(dR("Attribute")) = 4 Then
            strConnect = CStr(dR("Expression"))
            Exit For
        End If
    Next dR

    ' Step 5: Format SQL (skip formatting for pass-through queries)
    If Len(strConnect) = 0 Then
        Perf.OperationStart "Format SQL"
        With New clsSqlFormatter
            strSql = .FormatSQL(strSql, esdAccess)
        End With
        Perf.OperationEnd
    End If

    ' Step 6: Write .sql file
    Perf.OperationStart "Write SQL"
    VerifyPath strSqlFile
    WriteFile strSql, strSqlFile
    Perf.OperationEnd

    ' Step 7: Parse LvProp for properties and column metadata
    Perf.OperationStart "Parse LvProp"
    Set dJsonProps = New Dictionary
    Set dJsonColumns = New Dictionary
    If IsArrayAllocated(bteLvProp) Then
        Dim cPropParser As New clsLvPropParser
        Set dProps = cPropParser.ParseLvProp(bteLvProp, Options.SanitizeLevel)
        ' Extract table-level properties (query properties)
        If dProps.Exists("TableProperties") Then
            Set dJsonProps = dProps("TableProperties")
        End If
        ' Extract field-level properties (column metadata)
        If dProps.Exists("FieldProperties") Then
            Set dFieldProps = dProps("FieldProperties")
            ' Build sorted column dictionary
            Dim dSorted As Dictionary
            Set dSorted = SortDictionaryByKeys(dFieldProps)
            Dim dColEntry As Dictionary
            Dim dFieldPropSet As Dictionary
            Dim varPropName As Variant
            For Each varKey In dSorted.Keys
                Set dColEntry = New Dictionary
                Set dFieldPropSet = dSorted(varKey)
                For Each varPropName In dFieldPropSet.Keys
                    If IsKnownColumnProperty(CStr(varPropName)) Then
                        dColEntry.Add CStr(varPropName), dFieldPropSet(varPropName)("Value")
                    Else
                        Dim dTyped As Dictionary
                        Set dTyped = New Dictionary
                        dTyped.Add "Type", DaoTypeToQdefPrefix(CLng(dFieldPropSet(varPropName)("Type")))
                        dTyped.Add "Value", dFieldPropSet(varPropName)("Value")
                        dColEntry.Add CStr(varPropName), dTyped
                    End If
                Next varPropName
                If Not dColEntry.Exists("AggregateType") Then
                    dColEntry.Add "AggregateType", CLng(-1)
                End If
                dJsonColumns.Add CStr(varKey), dColEntry
            Next varKey
        End If
    End If
    Perf.OperationEnd

    ' Step 8: Parse LvExtra for design layout (if Design View)
    Set dLayout = Nothing
    If blnHasLvExtra Then
        Perf.OperationStart "Parse LvExtra"
        Dim cExtraParser As New clsLvExtraParser
        Set dLayout = cExtraParser.ParseLvExtra(bteLvExtra)
        Perf.OperationEnd
    End If

    ' Step 9: Build .json file
    Perf.OperationStart "Write JSON"
    Set dItems = New Dictionary

    ' Query type
    dItems.Add "QueryType", lngQueryType

    ' Option flag (Attribute 3 Flag bitmask: DISTINCT, TOP, etc.)
    Dim lngOptionFlag As Long
    lngOptionFlag = 0
    For Each dR In colRows
        If CLng(dR("Attribute")) = 3 Then
            lngOptionFlag = CLng(dR("Flag"))
            Exit For
        End If
    Next dR
    If lngOptionFlag <> 0 Then
        dItems.Add "OptionFlag", lngOptionFlag
    End If

    ' Connection string (with env: substitution)
    If Len(strConnect) > 0 Then
        strEnvKey = ShouldUseEnvForConnection(strConnect)
        If Len(strEnvKey) > 0 Then
            dItems.Add "Connect", "env:" & strEnvKey
        Else
            dItems.Add "Connect", strConnect
        End If
    End If

    ' Query properties from LvProp (ODBCTimeout, UseTransaction, etc.)
    If dJsonProps.Count > 0 Then
        dItems.Add "QueryProperties", dJsonProps
    End If

    ' Column metadata (sorted for determinism)
    If dJsonColumns.Count > 0 Then
        dItems.Add "Columns", dJsonColumns
    End If

    ' Design layout (only for Design View queries)
    If Not dLayout Is Nothing Then
        If dLayout.Count > 0 Then
            dItems.Add "DesignLayout", ConvertLayoutForJson(dLayout)
        End If
    End If

    ' Collect document properties (Description, Hidden)
    CollectObjectMetadata dItems, "Tables", strQueryName, acQuery

    ' Build the full JSON structure
    Set dHeader = New Dictionary
    dHeader.Add "Class", "clsDbQuery"
    dHeader.Add "Description", strQueryName

    Set dFile = New Dictionary
    dFile.Add "Info", dHeader
    dFile.Add "Items", dItems

    WriteFile ConvertToJson(dFile, JSON_WHITESPACE), strJsonFile
    Perf.OperationEnd

    Log.Add "  " & strQueryName, Options.ShowDebug

    ' When ShowDebug is on, save native SaveAsText output as .qdf for comparison,
    ' and the COM .SQL property as .com.sql for comparison with our reconstructed .sql
    If Options.ShowDebug Then
        SaveComponentAsText acQuery, strQueryName, _
            Left$(strSqlFile, Len(strSqlFile) - 4) & ".qdf"

        Dim strComSql As String
        LogUnhandledErrors
        On Error Resume Next
        strComSql = dbs.QueryDefs(strQueryName).SQL
        On Error GoTo 0
        If Len(strComSql) > 0 Then
            If Len(strConnect) = 0 Then
                With New clsSqlFormatter
                    strComSql = .FormatSQL(strComSql, esdAccess)
                End With
            End If
            WriteFile strComSql, Left$(strSqlFile, Len(strSqlFile) - 4) & ".com.sql"
        End If
    End If

    ' Clean up any legacy .qdef or .bas files
    CleanupLegacyFiles strSqlFile

CleanUp:
    ' Cache the real error before any cleanup-induced errors clobber it.
    ' (DAO has no Recordset.State -- calling Close on an already-closed rst
    '  raises 3420, which would otherwise mask the real error or be logged
    '  as a phantom failure when the body actually succeeded.)
    Dim lngOrigErr As Long
    Dim strOrigDesc As String
    lngOrigErr = Err.Number
    strOrigDesc = Err.Description

    On Error Resume Next
    If Not rst Is Nothing Then rst.Close
    Set rst = Nothing
    Err.Clear

    ' Re-raise the original error (if any) so CatchAny logs the real cause
    If lngOrigErr <> 0 Then Err.Raise lngOrigErr, , strOrigDesc
    CatchAny eelError, "Error exporting query '" & m_Query.Name & "'", _
        ModuleName(Me) & ".ExportNewFormat", True, True, True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ConvertLayoutForJson
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Convert the LvExtra parsed layout (with Collection of tables) into
'           : a JSON-serializable Dictionary structure.
'---------------------------------------------------------------------------------------
'
Private Function ConvertLayoutForJson(dLayout As Dictionary) As Dictionary

    Dim dResult As New Dictionary
    Dim dTbl As Dictionary
    Dim colTables As Collection
    Dim colJsonTables As Collection
    Dim dJsonTbl As Dictionary

    ' Copy scalar values
    If dLayout.Exists("state") Then dResult.Add "State", dLayout("state")
    If dLayout.Exists("window") Then dResult.Add "Window", dLayout("window")
    If dLayout.Exists("designerPane") Then dResult.Add "DesignerPane", dLayout("designerPane")
    If dLayout.Exists("gridLeft") Then dResult.Add "GridLeft", dLayout("gridLeft")
    If dLayout.Exists("gridTop") Then dResult.Add "GridTop", dLayout("gridTop")
    If dLayout.Exists("columnsShown") Then dResult.Add "ColumnsShown", dLayout("columnsShown")

    ' Convert table Collection to JSON-serializable array-like structure
    If dLayout.Exists("tables") Then
        Set colTables = dLayout("tables")
        Set colJsonTables = New Collection
        For Each dTbl In colTables
            Set dJsonTbl = New Dictionary
            dJsonTbl.Add "Name", CStr(dTbl("name"))
            dJsonTbl.Add "Alias", Nz(dTbl("alias"), vbNullString)
            dJsonTbl.Add "Left", CLng(dTbl("left"))
            dJsonTbl.Add "Top", CLng(dTbl("top"))
            dJsonTbl.Add "Right", CLng(dTbl("right"))
            dJsonTbl.Add "Bottom", CLng(dTbl("bottom"))
            dJsonTbl.Add "ScrollTop", CLng(dTbl("scrollTop"))
            colJsonTables.Add dJsonTbl
        Next dTbl
        dResult.Add "Tables", colJsonTables
    End If

    Set ConvertLayoutForJson = dResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : CleanupLegacyFiles
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Remove legacy .qdef and .bas files that are no longer needed after
'           : migrating to .sql + .json format.
'---------------------------------------------------------------------------------------
'
Private Sub CleanupLegacyFiles(strSqlFile As String)
    Dim strBase As String
    strBase = Left$(strSqlFile, Len(strSqlFile) - 4)
    If FSO.FileExists(strBase & ".qdef") Then DeleteFile strBase & ".qdef"
    If FSO.FileExists(strBase & ".bas") Then DeleteFile strBase & ".bas"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportLegacy
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Legacy export using SaveAsText, used when format version < 5.0.
'---------------------------------------------------------------------------------------
'
Private Sub ExportLegacy(strFile As String, Optional blnAlternate As Boolean = False)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim strHash As String
    Dim strSql As String
    Dim strSqlFile As String
    Dim dbs As DAO.Database

    strHash = SaveComponentAsText(acQuery, m_Query.Name, strFile)

    If blnAlternate Then
        VCSIndex.Update Me, eatAltExport, strHash, _
            strMetaHash:=GetMetadataHash("Tables", m_Query.Name, acQuery)
    Else
        If Options.SaveQuerySQL Then
            Perf.OperationStart "Save Query SQL"
            Set dbs = SharedDb
            strSqlFile = IDbComponent_BaseFolder & GetSafeFileName(m_Query.Name) & ".sql"
            LogUnhandledErrors
            On Error Resume Next
            strSql = dbs.QueryDefs(m_Query.Name).SQL
            If Catch(3258) Then
                Log.Error eelWarning, "Unable to access SQL property of query '" & _
                    m_Query.Name & "'. SQL view may not be available for this query.", _
                    ModuleName(Me) & ".Export"
            ElseIf Err.Number Then
                Log.Error eelError, "Unexpected error accessing SQL property of query '" & _
                    m_Query.Name & "'. Error " & Err.Number & ": " & Err.Description, _
                    ModuleName(Me) & ".Export"
            End If
            ' Restore outer error mode (was set at top of sub) instead of forcing GoTo 0,
            ' which would defeat the wrap in release mode.
            If DebugMode(False) Then On Error GoTo 0 Else On Error Resume Next
            If strSql <> vbNullString Then
                Select Case dbs.QueryDefs(m_Query.Name).Type
                    Case dbQSQLPassThrough, dbQSPTBulk
                    Case Else
                        With New clsSqlFormatter
                            Perf.OperationStart "Format SQL"
                            strSql = .FormatSQL(strSql, esdAccess)
                            Perf.OperationEnd
                        End With
                End Select
                WriteFile strSql, strSqlFile
            End If
            Perf.OperationEnd
            Log.Add "  " & m_Query.Name & " (SQL)", Options.ShowDebug
        End If
        VCSIndex.Update Me, eatExport, GetFileHash(strFile), _
            strMetaHash:=GetMetadataHash("Tables", m_Query.Name, acQuery)
    End If

    CatchAny eelError, "Error exporting query (legacy) '" & m_Query.Name & "'", _
        ModuleName(Me) & ".ExportLegacy"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller / Indigo
' Date      : 10/24/2020
' Purpose   : Import the individual database component from a file.
'           :
'           : v5.0+: Reads .sql + .json, generates .qdef on-the-fly, calls LoadFromText.
'           : Legacy: Still supports importing .qdef and .bas files.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Route to appropriate import method based on file extension
    If strFile Like "*.sql" Then
        ImportNewFormat strFile
    ElseIf strFile Like "*.qdef" Or strFile Like "*.bas" Then
        ImportLegacyFormat strFile
    End If

    CatchAny eelError, "Error importing query from '" & strFile & "'", _
        ModuleName(Me) & ".IDbComponent_Import"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportNewFormat
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Import from .sql + .json files by generating a .qdef on-the-fly
'           : and calling LoadFromText.
'---------------------------------------------------------------------------------------
'
Private Sub ImportNewFormat(strSqlFile As String)

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Dim strQueryName As String
    Dim strSql As String
    Dim strJsonFile As String
    Dim strQdefContent As String
    Dim strTempFile As String
    Dim dFile As Dictionary
    Dim dItems As Dictionary
    Dim dProperties As Dictionary
    Dim dColumns As Dictionary
    Dim dLayout As Dictionary
    Dim cComposer As clsQueryComposer
    Dim blnDesignView As Boolean
    Dim blnIsPassThrough As Boolean
    Dim strConnect As String
    Dim lngOptionFlag As Long

    strQueryName = GetObjectNameFromFileName(strSqlFile)

    ' Read .sql file
    strSql = ReadFile(strSqlFile)
    If Len(strSql) = 0 Then
        Log.Error eelError, "Empty SQL file for query '" & strQueryName & "'.", _
            "clsDbQuery.ImportNewFormat"
        GoTo CleanUp
    End If

    ' Read .json file
    strJsonFile = SwapExtension(strSqlFile, "json")
    Set dProperties = New Dictionary
    Set dColumns = New Dictionary
    Set dLayout = Nothing

    If FSO.FileExists(strJsonFile) Then
        Set dFile = ReadJsonFile(strJsonFile)
        If Not dFile Is Nothing Then
            If dFile.Exists("Items") Then
                Set dItems = dFile("Items")

                ' Option flag bitmask (DISTINCT, TOP, etc.)
                If dItems.Exists("OptionFlag") Then
                    lngOptionFlag = CLng(dItems("OptionFlag"))
                End If

                ' Connection string
                If dItems.Exists("Connect") Then
                    strConnect = CStr(dItems("Connect"))
                    ' Resolve env: references
                    If Left$(strConnect, 4) = "env:" Then
                        strConnect = ResolveEnvConnection(strConnect)
                    End If
                    blnIsPassThrough = True
                End If

                ' Query properties (from LvProp)
                If dItems.Exists("QueryProperties") Then
                    Set dProperties = dItems("QueryProperties")
                End If

                ' Columns
                If dItems.Exists("Columns") Then
                    Set dColumns = dItems("Columns")
                End If

                ' Design layout
                If dItems.Exists("DesignLayout") Then
                    Set dLayout = ConvertLayoutFromJson(dItems("DesignLayout"))
                End If
            End If
        End If
    End If

    ' Decompose SQL to determine if Design View is possible
    Set cComposer = New clsQueryComposer
    cComposer.DecomposeSQL strSql, blnIsPassThrough
    cComposer.RawSQL = strSql
    If blnIsPassThrough Then cComposer.ConnectString = strConnect

    ' Determine view mode
    blnDesignView = cComposer.IsDesignerCompatible And Not dLayout Is Nothing

    ' Generate .qdef content
    strQdefContent = cComposer.GenerateQdef(blnDesignView, dProperties, dColumns, dLayout, lngOptionFlag)

    ' Write to temp file and import
    strTempFile = GetTempFile
    WriteFile strQdefContent, strTempFile

    ' Attempt to import (suppress errors on Design View so the fallback
    ' to SQL View can happen without showing a dialog for each query)
    If LoadComponentFromText(acQuery, strQueryName, strTempFile, blnDesignView) = False Then
        ' Design View import failed -- fall back to SQL View
        If blnDesignView Then
            Log.Error eelWarning, "Design View import failed for '" & strQueryName & _
                "'; falling back to SQL View (layout will be lost).", _
                "clsDbQuery.ImportNewFormat"
            ' Preserve the failed Design View .qdef for debugging
            If Options.ShowDebug Then
                Dim strFailedFile As String
                strFailedFile = IDbComponent_BaseFolder & GetSafeFileName(strQueryName) & ".failed.tmp"
                If FSO.FileExists(strFailedFile) Then DeleteFile strFailedFile
                FSO.CopyFile strTempFile, strFailedFile
            End If
            strQdefContent = cComposer.GenerateQdef(False, dProperties, dColumns, , lngOptionFlag)
            WriteFile strQdefContent, strTempFile
            If LoadComponentFromText(acQuery, strQueryName, strTempFile) = False Then
                Log.Error eelError, "Failed to import query '" & strQueryName & "'.", _
                    "clsDbQuery.ImportNewFormat"
                GoTo CleanUp
            End If
        Else
            Log.Error eelError, "Failed to import query '" & strQueryName & "'.", _
                "clsDbQuery.ImportNewFormat"
            GoTo CleanUp
        End If
    End If

    Set m_Query = CurrentData.AllQueries(strQueryName)

    ' Apply document metadata from the .json file
    ImportObjectMetadata strJsonFile, "Tables", strQueryName, acQuery

    ' Update index
    VCSIndex.Update Me, eatImport, GetFileHash(strSqlFile), _
        strMetaHash:=GetMetadataHash("Tables", strQueryName, acQuery)

CleanUp:
    ' Log any unexpected error first (before file ops below can clobber Err)
    CatchAny eelError, "Error importing query '" & strQueryName & "'", _
        ModuleName(Me) & ".ImportNewFormat", True, True, True

    On Error Resume Next
    If Len(strTempFile) > 0 Then
        If FSO.FileExists(strTempFile) Then
            ' When ShowDebug is on, preserve the generated .qdef as .tmp for inspection
            If Options.ShowDebug Then
                Dim strDebugFile As String
                strDebugFile = IDbComponent_BaseFolder & GetSafeFileName(strQueryName) & ".tmp"
                If FSO.FileExists(strDebugFile) Then DeleteFile strDebugFile
                FSO.CopyFile strTempFile, strDebugFile
            End If
            DeleteFile strTempFile
        End If
    End If
    ' Don't leak cleanup-induced errors to the caller
    Err.Clear

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ConvertLayoutFromJson
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Convert the JSON layout structure back to the Dictionary format
'           : expected by clsQueryComposer.GenerateQdef.
'---------------------------------------------------------------------------------------
'
Private Function ConvertLayoutFromJson(dJsonLayout As Dictionary) As Dictionary

    Dim dResult As New Dictionary
    Dim colTables As New Collection
    Dim dTbl As Dictionary
    Dim dJsonTbl As Dictionary
    Dim varTbl As Variant

    ' Scalar values
    If dJsonLayout.Exists("State") Then dResult.Add "state", dJsonLayout("State")
    If dJsonLayout.Exists("Window") Then dResult.Add "window", dJsonLayout("Window")
    If dJsonLayout.Exists("DesignerPane") Then dResult.Add "designerPane", dJsonLayout("DesignerPane")
    If dJsonLayout.Exists("GridLeft") Then dResult.Add "gridLeft", dJsonLayout("GridLeft")
    If dJsonLayout.Exists("GridTop") Then dResult.Add "gridTop", dJsonLayout("GridTop")
    If dJsonLayout.Exists("ColumnsShown") Then dResult.Add "columnsShown", dJsonLayout("ColumnsShown")

    ' Tables
    If dJsonLayout.Exists("Tables") Then
        For Each varTbl In dJsonLayout("Tables")
            If TypeOf varTbl Is Dictionary Then
                Set dJsonTbl = varTbl
            Else
                GoTo NextTable
            End If
            Set dTbl = New Dictionary
            dTbl.Add "name", CStr(dJsonTbl("Name"))
            dTbl.Add "alias", Nz(dJsonTbl("Alias"), vbNullString)
            dTbl.Add "left", CLng(dJsonTbl("Left"))
            dTbl.Add "top", CLng(dJsonTbl("Top"))
            dTbl.Add "right", CLng(dJsonTbl("Right"))
            dTbl.Add "bottom", CLng(dJsonTbl("Bottom"))
            dTbl.Add "scrollTop", CLng(dJsonTbl("ScrollTop"))
            colTables.Add dTbl
NextTable:
        Next varTbl
        dResult.Add "tables", colTables
    End If

    Set ConvertLayoutFromJson = dResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : ImportLegacyFormat
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Import from legacy .qdef or .bas files (backward compatibility).
'---------------------------------------------------------------------------------------
'
Private Sub ImportLegacyFormat(strFile As String)

    Dim dbs As DAO.Database
    Dim strQueryName As String
    Dim strFileSql As String
    Dim strSql As String
    Dim blnRecreated As Boolean

    strQueryName = GetObjectNameFromFileName(strFile)

    If Options.ForceImportOriginalQuerySQL Then
        If LoadComponentFromText(acQuery, strQueryName, strFile, True) = False Then
            strFileSql = SwapExtension(strFile, "sql")
            strSql = ReadFile(strFileSql)
            If strSql <> vbNullString Then
                Set dbs = SharedDb
                dbs.CreateQueryDef strQueryName, strSql
                blnRecreated = True
                Log.Error eelWarning, "  Restored query " & strQueryName & _
                    " using SQL. Some visual design properties may be lost.", _
                    "clsDbQuery.ImportLegacyFormat"
            Else
                Log.Error eelError, "  Couldn't get original SQL query for " & strQueryName
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    Else
        If LoadComponentFromText(acQuery, strQueryName, strFile) = False Then
            Exit Sub
        End If
    End If

    Set m_Query = CurrentData.AllQueries(strQueryName)
    ImportObjectMetadata SwapExtension(strFile, "json"), "Tables", strQueryName, acQuery

    If Options.ForceImportOriginalQuerySQL And (blnRecreated = False) Then
        strFileSql = SwapExtension(strFile, "sql")
        strSql = ReadFile(strFileSql)
        If strSql <> vbNullString Then
            Set dbs = SharedDb
            dbs.QueryDefs(strQueryName).SQL = strSql
            Log.Add "  Restored original SQL for " & strQueryName, Options.ShowDebug
        Else
            Log.Add "  Couldn't get original SQL query for " & strQueryName
        End If
    End If

    VCSIndex.Update Me, eatImport, GetFileHash(strFile), _
        strMetaHash:=GetMetadataHash("Tables", strQueryName, acQuery)

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
    DeleteObjectIfExists acQuery, GetObjectNameFromFileName(strFile)
    If FSO.FileExists(strFile) Then
        IDbComponent_Import strFile
    Else
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
    MoveFileIfExists strFromFolder & FSO.GetBaseName(IDbComponent_SourceFile) & ".json", strToFolder
    ' Also move legacy files if present
    MoveFileIfExists strFromFolder & FSO.GetBaseName(IDbComponent_SourceFile) & ".qdef", strToFolder
    MoveFileIfExists strFromFolder & FSO.GetBaseName(IDbComponent_SourceFile) & ".bas", strToFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary

    Dim qry As AccessObject
    Dim cQuery As IDbComponent
    Dim blnNeedAll As Boolean

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        blnNeedAll = (m_Items(False) Is Nothing)
        If blnNeedAll Then Set m_Items(False) = New Dictionary
        If blnModifiedOnly Then Set m_Items(True) = New Dictionary
        For Each qry In CurrentData.AllQueries
            Set cQuery = New clsDbQuery
            Set cQuery.DbObject = qry
            If blnNeedAll Then m_Items(False).Add cQuery.SourceFile, cQuery
            If blnModifiedOnly Then
                If cQuery.IsModified Then m_Items(True).Add cQuery.SourceFile, cQuery
            End If
        Next qry
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_Items(blnModifiedOnly)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'           : New format (.sql) takes priority; falls back to legacy .qdef/.bas
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Dictionary
    If m_FileList Is Nothing Then
        Dim dNewFormat As Dictionary
        Dim varKey As Variant
        Dim strQdefEquiv As String
        Dim strBasEquiv As String
        Dim strJsonEquiv As String

        ' Legacy format first: .qdef and .bas files always take priority
        Set m_FileList = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.qdef")
        MergeDictionary m_FileList, GetFilePathsInFolder(IDbComponent_BaseFolder, "*.bas")

        ' New format: .sql + .json pairs (only if no legacy .qdef/.bas exists)
        Set dNewFormat = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.sql")
        For Each varKey In dNewFormat.Keys
            strJsonEquiv = SwapExtension(CStr(varKey), "json")
            strQdefEquiv = SwapExtension(CStr(varKey), "qdef")
            strBasEquiv = SwapExtension(CStr(varKey), "bas")
            If FSO.FileExists(strJsonEquiv) _
                And Not m_FileList.Exists(strQdefEquiv) _
                And Not m_FileList.Exists(strBasEquiv) Then
                m_FileList.Add CStr(varKey), dNewFormat(varKey)
            End If
        Next varKey
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
    IDbComponent_IsModified = (DateTruncToSeconds(m_Query.DateModified) <> VCSIndex.Item(Me).ObjectDate)

    ' Check metadata (Description/Hidden) which don't update DateModified
    If Not IDbComponent_IsModified Then
        IDbComponent_IsModified = VCSIndex.Item(Me).MetaHash <> _
            GetMetadataHash("Tables", m_Query.Name, acQuery)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : DateModified
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_DateModified() As Date
    If m_Query Is Nothing Then Exit Function
    IDbComponent_DateModified = m_Query.DateModified
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "Queries"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "queries" & PathSep
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
    IDbComponent_FileExtensions.Add "sql"
    IDbComponent_FileExtensions.Add "json"
    ' Legacy extensions (for backward compatibility)
    IDbComponent_FileExtensions.Add "qdef"
    IDbComponent_FileExtensions.Add "bas"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    If m_Query Is Nothing Then Exit Property
    IDbComponent_Name = m_Query.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'           : v5.0+: Uses .sql as the primary source file.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    If m_Query Is Nothing Then Exit Property
    If Options.ExportFormatVersion >= EFV_5_0_0 Then
        IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Query.Name) & ".sql"
    Else
        IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Query.Name) & ".qdef"
    End If
End Property


Private Property Get IDbComponent_Count(Optional blnModifiedOnly As Boolean = False) As Long
    IDbComponent_Count = IDbComponent_GetAllFromDB(blnModifiedOnly).Count
End Property

Private Property Get IDbComponent_QuickCount() As Long
    Static lngCount As Long
    If lngCount = 0 Then lngCount = CurrentData.AllQueries.Count - 1
    IDbComponent_QuickCount = lngCount + 1
End Property

Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbQuery
End Property

Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_Query
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Query = RHS
End Property

Private Property Get IDbComponent_SingleFile() As Boolean
    IDbComponent_SingleFile = False
End Property

Public Property Get Parent() As IDbComponent
    Set Parent = Me
End Property


'---------------------------------------------------------------------------------------
' Procedure : EscapeSqlQuote
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Escape double quotes in a string for use in an Access SQL query.
'---------------------------------------------------------------------------------------
'
Private Function EscapeSqlQuote(strValue As String) As String
    EscapeSqlQuote = Replace(strValue, """", """""")
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsArrayAllocated
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Returns True if a byte array has been assigned data.
'---------------------------------------------------------------------------------------
'
Private Function IsArrayAllocated(bteArr() As Byte) As Boolean
    On Error Resume Next
    IsArrayAllocated = (UBound(bteArr) >= 0)
    On Error GoTo 0
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsKnownColumnProperty
' Author    : Adam Waller
' Date      : 4/13/2026
' Purpose   : Returns True if the property name is a known column property whose
'           : QDEF type can be inferred from the name (no explicit type needed in JSON).
'---------------------------------------------------------------------------------------
'
Private Function IsKnownColumnProperty(strName As String) As Boolean
    Select Case strName
        Case "AggregateType", "ColumnWidth", "ColumnHidden", "ColumnOrder", _
             "Caption", "Description", "TextAlign", "DisplayControl", _
             "ResultType", "CurrencyLCID", "ShowDatePicker", _
             "IMEMode", "IMESentenceMode"
            IsKnownColumnProperty = True
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : DaoTypeToQdefPrefix
' Author    : Adam Waller
' Date      : 4/13/2026
' Purpose   : Maps a DAO data type code to the corresponding QDEF type prefix string.
'           : Used when storing custom (unknown) properties with explicit type info.
'---------------------------------------------------------------------------------------
'
Private Function DaoTypeToQdefPrefix(lngType As Long) As String
    Select Case lngType
        Case 1:  DaoTypeToQdefPrefix = "dbBoolean"
        Case 2:  DaoTypeToQdefPrefix = "dbByte"
        Case 3:  DaoTypeToQdefPrefix = "dbInteger"
        Case 4:  DaoTypeToQdefPrefix = "dbLong"
        Case 5:  DaoTypeToQdefPrefix = "dbCurrency"
        Case 6:  DaoTypeToQdefPrefix = "dbSingle"
        Case 7:  DaoTypeToQdefPrefix = "dbDouble"
        Case 8:  DaoTypeToQdefPrefix = "dbDate"
        Case 10: DaoTypeToQdefPrefix = "dbText"
        Case 12: DaoTypeToQdefPrefix = "dbMemo"
        Case Else: DaoTypeToQdefPrefix = "dbMemo"
    End Select
End Function
