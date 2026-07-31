Attribute VB_Name = "modTestTableData"
'---------------------------------------------------------------------------------------
' Module    : modTestTableData
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Unit and integration tests for deterministic table data export ordering.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Components")
'@Tag("unit")


Private Const TEST_TABLE_ORDER As String = "vcs_test_tabledata_order"
Private Const TEST_TABLE_TEXT_ORDER As String = "vcs_test_tabledata_text"
Private Const TEST_TABLE_PLAIN As String = "vcs_test_plain"
Private Const TEST_TABLE_OLE As String = "vcs_test_ole"
Private Const TEST_TABLE_SORTFIELDS As String = "vcs_test_sortfields"
Private Const TEST_TABLE_PRIMARY_KEY As String = "vcs_test_sortfields_pk"
Private Const TEST_TABLE_BINARY As String = "vcs_test_binary_pk"
Private Const TEST_TABLE_MERGE As String = "vcs_test_merge"
Private Const TEST_TABLE_MERGE_COMPOSITE As String = "vcs_test_merge_composite"
Private Const TEST_TABLE_MERGE_NULLS As String = "vcs_test_merge_nulls"
Private Const TEST_TABLE_MERGE_PARENT As String = "vcs_test_merge_parent"
Private Const TEST_TABLE_MERGE_CHILD As String = "vcs_test_merge_child"

' The XML merge table deliberately carries a field with the same name as the table.
Private Const TEST_TABLE_MERGE_XML As String = "vcs_test_merge_xml"

' Wide enough that the reconcile cannot assign and compare every field in one statement.
Private Const TEST_TABLE_MERGE_WIDE As String = "vcs_test_merge_wide"
Private Const WIDE_FIELD_COUNT As Long = 80


Public Sub TestEscapeXmlName()
    TestAssert EscapeXmlName("NotReq'd") = "NotReq_x0027_d", "apostrophe"
    TestAssert EscapeXmlName("Please" & Chr$(34) & "d" & Chr$(34) & "don" & Chr$(39) & "t" & Chr$(34) & "use") = _
        "Please_x0022_d_x0022_don_x0027_t_x0022_use", "quotes and apostrophes"
    TestAssert EscapeXmlName("ID") = "ID", "simple name unchanged"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestEscapeXmlNameUnderscore
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Access escapes an underscore only when a lowercase "x" follows it, so the
'           : pair cannot be read as the start of an _xHHHH_ sequence. Expectations here
'           : were captured from actual Application.ExportXML output.
'---------------------------------------------------------------------------------------
'
Public Sub TestEscapeXmlNameUnderscore()
    TestAssert EscapeXmlName("a_b") = "a_b", "underscore before a letter is left alone"
    TestAssert EscapeXmlName("a_x1") = "a_x005F_x1", "underscore before lowercase x is escaped"
    TestAssert EscapeXmlName("b_X2") = "b_X2", "underscore before uppercase X is left alone"
    TestAssert EscapeXmlName("a_x005F_b") = "a_x005F_x005F_b", "existing escape sequence is re-escaped"
    TestAssert EscapeXmlName("a_xZZZZ_b") = "a_x005F_xZZZZ_b", "escaped even when the hex digits are invalid"
    TestAssert EscapeXmlName("a__x") = "a__x005F_x", "only the underscore adjacent to x is escaped"
    TestAssert EscapeXmlName("trail_") = "trail_", "trailing underscore is left alone"
End Sub


Public Sub TestTableRequiresXmlSchema()
    '@Tag("integration")

    Dim dbs As DAO.Database

    Set dbs = CurrentDb

    CreateTestTable dbs, TEST_TABLE_PLAIN, _
        "CREATE TABLE [" & TEST_TABLE_PLAIN & "] (Alpha TEXT(10), Beta LONG)"
    TestAssert Not TableRequiresXmlSchema(dbs.TableDefs(TEST_TABLE_PLAIN)), _
        "plain table does not need the embedded schema"
    DropTestTable TEST_TABLE_PLAIN, dbs

    CreateTestTable dbs, TEST_TABLE_OLE, _
        "CREATE TABLE [" & TEST_TABLE_OLE & "] (Alpha TEXT(10), Blob LONGBINARY)"
    TestAssert TableRequiresXmlSchema(dbs.TableDefs(TEST_TABLE_OLE)), _
        "OLE object field requires the embedded schema"
    DropTestTable TEST_TABLE_OLE, dbs

End Sub


Public Sub TestNormalizeNumericXmlSortValue_Order()
    Dim varValues(0 To 4) As String

    varValues(0) = NormalizeNumericXmlSortValue("-10")
    varValues(1) = NormalizeNumericXmlSortValue("-2")
    varValues(2) = NormalizeNumericXmlSortValue("0")
    varValues(3) = NormalizeNumericXmlSortValue("2")
    varValues(4) = NormalizeNumericXmlSortValue("10")

    TestAssert varValues(0) < varValues(1), "-10 before -2"
    TestAssert varValues(1) < varValues(2), "-2 before 0"
    TestAssert varValues(2) < varValues(3), "0 before 2"
    TestAssert varValues(3) < varValues(4), "2 before 10"
End Sub


Public Sub TestXmlSortKeyIgnoresValueLength()
    Dim astrValues(0 To 6) As String
    Dim astrKeys(0 To 6) As String
    Dim astrExpected(0 To 6) As String
    Dim lngIdx As Long
    Dim lngPos As Long

    astrValues(0) = "03.8"
    astrValues(1) = "03.81"
    astrValues(2) = "03.640"
    astrValues(3) = "48.4A"
    astrValues(4) = "64.6"
    astrValues(5) = "10.26int2"
    astrValues(6) = "03.94 (BETA)"

    For lngIdx = 0 To 6
        astrKeys(lngIdx) = BuildSingleTextXmlSortKey(astrValues(lngIdx), lngIdx)
    Next lngIdx

    QuickSortStringsBinary astrKeys

    For lngIdx = 0 To 6
        astrExpected(lngIdx) = astrValues(lngIdx)
    Next lngIdx
    QuickSortStringsBinary astrExpected

    For lngIdx = 0 To 6
        lngPos = XmlSortKeyOrdinal(astrKeys(lngIdx))
        TestAssert StrComp(astrValues(lngPos), astrExpected(lngIdx), vbBinaryCompare) = 0, _
            "key order matches text sort: " & astrExpected(lngIdx)
    Next lngIdx

    TestAssert StrComp(astrExpected(6), "64.6", vbBinaryCompare) = 0, "64.6 sorts last"
End Sub


Public Sub TestXmlSortKeyFieldBoundary()
    Dim colPartsA As New Collection
    Dim colPartsB As New Collection
    Dim strKeyA As String
    Dim strKeyB As String

    colPartsA.Add NormalizeXmlSortValue("ab", dbText)
    colPartsA.Add NormalizeXmlSortValue("z", dbText)
    colPartsB.Add NormalizeXmlSortValue("abc", dbText)
    colPartsB.Add NormalizeXmlSortValue("a", dbText)

    strKeyA = ComposeXmlSortKey(colPartsA, 0)
    strKeyB = ComposeXmlSortKey(colPartsB, 1)

    TestAssert StrComp(strKeyA, strKeyB, vbBinaryCompare) < 0, "(ab, z) before (abc, a)"
End Sub


Public Sub TestXmlSortKeyOrdinalRoundTrip()
    Dim colParts As New Collection
    Dim astrKeys(0 To 2) As String
    Dim lngIdx As Long

    colParts.Add NormalizeXmlSortValue("same", dbText)

    For lngIdx = 0 To 2
        astrKeys(lngIdx) = ComposeXmlSortKey(colParts, lngIdx)
    Next lngIdx

    QuickSortStringsBinary astrKeys

    For lngIdx = 0 To 2
        TestAssert XmlSortKeyOrdinal(astrKeys(lngIdx)) = lngIdx, "ordinal round-trip " & lngIdx
    Next lngIdx
End Sub


Public Sub TestGetTableSortFields_UsesPrimaryKey()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim dSort As Dictionary

    Set dbs = CurrentDb
    CreateTestTable dbs, TEST_TABLE_PRIMARY_KEY, _
        "CREATE TABLE [" & TEST_TABLE_PRIMARY_KEY & "] (ID LONG, ObjectType LONG, CONSTRAINT PK PRIMARY KEY (ID, ObjectType))"

    Set dSort = GetTableSortFields(dbs.TableDefs(TEST_TABLE_PRIMARY_KEY))
    TestAssert dSort.Count = 2, "composite primary key has two fields"
    TestAssert dSort.Exists("ID"), "ID in sort fields"
    TestAssert dSort.Exists("ObjectType"), "ObjectType in sort fields"

    DropTestTable TEST_TABLE_PRIMARY_KEY, dbs
End Sub


Public Sub TestGetTableSortFields_NoKeyUsesAllFields()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim dSort As Dictionary

    Set dbs = CurrentDb
    CreateTestTable dbs, TEST_TABLE_SORTFIELDS, _
        "CREATE TABLE [" & TEST_TABLE_SORTFIELDS & "] (Alpha TEXT(10), Beta LONG)"

    Set dSort = GetTableSortFields(dbs.TableDefs(TEST_TABLE_SORTFIELDS))
    TestAssert dSort.Count = 2, "both non-binary fields"
    TestAssert dSort.Exists("Alpha"), "Alpha"
    TestAssert dSort.Exists("Beta"), "Beta"

    DropTestTable TEST_TABLE_SORTFIELDS, dbs
End Sub


Public Sub TestTableDataExport_DeterministicRowOrder()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strTdfFile As String
    Dim strXmlFile As String
    Dim strContent As String

    Set dbs = CurrentDb
    CreateTestTable dbs, TEST_TABLE_ORDER, _
        "CREATE TABLE [" & TEST_TABLE_ORDER & "] (SortID LONG PRIMARY KEY, Label TEXT(50))"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_ORDER & "] (SortID, Label) VALUES (3, 'C')"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_ORDER & "] (SortID, Label) VALUES (1, 'A')"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_ORDER & "] (SortID, Label) VALUES (2, 'B')"

    strTdfFile = GetTempFile & ".txt"
    ExportTestTableData TEST_TABLE_ORDER, etdTabDelimited, strTdfFile
    strContent = ReadFile(strTdfFile)
    TestAssert InStr(strContent, "1" & vbTab & "A") > InStr(strContent, "SortID"), "TDF row 1 present"
    TestAssert InStr(strContent, "2" & vbTab & "B") > InStr(strContent, "1" & vbTab & "A"), "TDF row 2 after 1"
    TestAssert InStr(strContent, "3" & vbTab & "C") > InStr(strContent, "2" & vbTab & "B"), "TDF row 3 after 2"
    DeleteFile strTdfFile

    strXmlFile = GetTempFile & ".xml"
    ExportTestTableData TEST_TABLE_ORDER, etdXML, strXmlFile
    strContent = ReadFile(strXmlFile)
    TestAssert InStr(strContent, "<SortID>1</SortID>") < InStr(strContent, "<SortID>2</SortID>"), "XML ID 1 before 2"
    TestAssert InStr(strContent, "<SortID>2</SortID>") < InStr(strContent, "<SortID>3</SortID>"), "XML ID 2 before 3"
    DeleteFile strXmlFile

    CreateTestTable dbs, TEST_TABLE_TEXT_ORDER, _
        "CREATE TABLE [" & TEST_TABLE_TEXT_ORDER & "] (VersionNo TEXT(50) PRIMARY KEY)"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_TEXT_ORDER & "] (VersionNo) VALUES ('03.81')"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_TEXT_ORDER & "] (VersionNo) VALUES ('64.6')"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_TEXT_ORDER & "] (VersionNo) VALUES ('03.8')"

    strXmlFile = GetTempFile & ".xml"
    ExportTestTableData TEST_TABLE_TEXT_ORDER, etdXML, strXmlFile
    strContent = ReadFile(strXmlFile)
    TestAssert InStr(strContent, "<VersionNo>03.8</VersionNo>") < InStr(strContent, "<VersionNo>03.81</VersionNo>"), _
        "XML text PK 03.8 before 03.81"
    TestAssert InStr(strContent, "<VersionNo>03.81</VersionNo>") < InStr(strContent, "<VersionNo>64.6</VersionNo>"), _
        "XML text PK 03.81 before 64.6"
    DeleteFile strXmlFile

    ' The sorted XML export routes through a temporary query, which must not survive.
    TestAssert DCount("*", "MSysObjects", "Name Like 'vcs_tmp_sort_export*'") = 0, _
        "temporary sort query removed after export"

    DropTestTable TEST_TABLE_TEXT_ORDER, dbs
    DropTestTable TEST_TABLE_ORDER, dbs
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableDataSortQuerySqlReuse
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Verify Application.ExportXML honors a reassigned .SQL on the reused
'           : temporary sort query without QueryDefs.Refresh.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDataSortQuerySqlReuse()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strQuery As String
    Dim strFile As String
    Dim strContent As String

    Set dbs = CurrentDb
    CreateTestTable dbs, TEST_TABLE_ORDER, _
        "CREATE TABLE [" & TEST_TABLE_ORDER & "] (SortID LONG PRIMARY KEY, Label TEXT(50))"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_ORDER & "] (SortID, Label) VALUES (1, 'A')"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_ORDER & "] (SortID, Label) VALUES (2, 'B')"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_ORDER & "] (SortID, Label) VALUES (3, 'C')"

    PrepareTableDataSortExport
    strQuery = AssignTableDataSortQuery("SELECT * FROM [" & TEST_TABLE_ORDER & "] ORDER BY [SortID]")
    TestAssert Len(strQuery) > 0, "temporary sort query created"

    strFile = GetTempFile & ".xml"
    Application.ExportXML acExportQuery, strQuery, strFile
    strContent = ReadFile(strFile)
    TestAssert InStr(strContent, "<SortID>1</SortID>") < InStr(strContent, "<SortID>3</SortID>"), _
        "ascending order on first SQL"

    strQuery = AssignTableDataSortQuery("SELECT * FROM [" & TEST_TABLE_ORDER & "] ORDER BY [SortID] DESC")
    Application.ExportXML acExportQuery, strQuery, strFile
    strContent = ReadFile(strFile)
    TestAssert InStr(strContent, "<SortID>3</SortID>") < InStr(strContent, "<SortID>1</SortID>"), _
        "descending order after SQL reuse without QueryDefs.Refresh"

    DeleteFile strFile
    ReleaseTableDataSortExport
    DropTestTable TEST_TABLE_ORDER, dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableDataSortQueryHoldsRequestedSql
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : The reused query must never be handed back still holding a prior table's
'           : SQL, which would export that table's rows into this table's source file.
'           : Asserted on the saved SQL rather than on a rejected statement, because the
'           : engine defers table-name resolution and accepts SQL naming a missing table.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDataSortQueryHoldsRequestedSql()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strQuery As String
    Dim strSaved As String

    Set dbs = CurrentDb
    CreateTestTable dbs, TEST_TABLE_ORDER, _
        "CREATE TABLE [" & TEST_TABLE_ORDER & "] (SortID LONG PRIMARY KEY, Label TEXT(50))"
    CreateTestTable dbs, TEST_TABLE_TEXT_ORDER, _
        "CREATE TABLE [" & TEST_TABLE_TEXT_ORDER & "] (VersionNo TEXT(50) PRIMARY KEY)"

    PrepareTableDataSortExport

    strQuery = AssignTableDataSortQuery("SELECT * FROM [" & TEST_TABLE_ORDER & "] ORDER BY [SortID]")
    TestAssert Len(strQuery) > 0, "query created for the first table"
    strSaved = SharedDb.QueryDefs(strQuery).SQL
    TestAssert InStr(1, strSaved, TEST_TABLE_ORDER, vbTextCompare) > 0, _
        "query holds the first table's SQL"

    strQuery = AssignTableDataSortQuery("SELECT * FROM [" & TEST_TABLE_TEXT_ORDER & "] ORDER BY [VersionNo]")
    TestAssert Len(strQuery) > 0, "query repointed for the second table"
    strSaved = SharedDb.QueryDefs(strQuery).SQL
    TestAssert InStr(1, strSaved, TEST_TABLE_TEXT_ORDER, vbTextCompare) > 0, _
        "query holds the second table's SQL"
    TestAssert InStr(1, strSaved, TEST_TABLE_ORDER, vbTextCompare) = 0, _
        "no trace of the first table's SQL remains"

    ReleaseTableDataSortExport
    TestAssert DCount("*", "MSysObjects", "Name Like 'vcs_tmp_sort_export*'") = 0, _
        "no temporary sort query left behind"

    DropTestTable TEST_TABLE_TEXT_ORDER, dbs
    DropTestTable TEST_TABLE_ORDER, dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableDataSortQueryRecoversWhenDeleted
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : The query is reused across a whole export, so losing it partway through
'           : (interrupted run, external cleanup) must not stop the remaining tables from
'           : exporting through the sorted path.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDataSortQueryRecoversWhenDeleted()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strSql As String
    Dim strFirst As String
    Dim strSecond As String

    Set dbs = CurrentDb
    CreateTestTable dbs, TEST_TABLE_ORDER, _
        "CREATE TABLE [" & TEST_TABLE_ORDER & "] (SortID LONG PRIMARY KEY, Label TEXT(50))"
    strSql = "SELECT * FROM [" & TEST_TABLE_ORDER & "] ORDER BY [SortID]"

    PrepareTableDataSortExport
    strFirst = AssignTableDataSortQuery(strSql)
    TestAssert Len(strFirst) > 0, "query created"

    SharedDb.QueryDefs.Delete strFirst

    strSecond = AssignTableDataSortQuery(strSql)
    TestAssert Len(strSecond) > 0, "query recreated after being deleted externally"
    TestAssert InStr(1, SharedDb.QueryDefs(strSecond).SQL, TEST_TABLE_ORDER, vbTextCompare) > 0, _
        "recreated query holds the requested SQL"

    ReleaseTableDataSortExport
    TestAssert DCount("*", "MSysObjects", "Name Like 'vcs_tmp_sort_export*'") = 0, _
        "no temporary sort query left behind"

    DropTestTable TEST_TABLE_ORDER, dbs

End Sub


Public Sub TestGetTableMergeKey_RequiresUniqueIndex()
    '@Tag("integration")

    Dim dbs As DAO.Database

    Set dbs = CurrentDb

    CreateTestTable dbs, TEST_TABLE_PRIMARY_KEY, _
        "CREATE TABLE [" & TEST_TABLE_PRIMARY_KEY & "] (ID LONG, ObjectType LONG, CONSTRAINT PK PRIMARY KEY (ID, ObjectType))"
    TestAssert GetTableMergeKey(dbs.TableDefs(TEST_TABLE_PRIMARY_KEY)).Count = 2, _
        "composite primary key returns both fields"
    DropTestTable TEST_TABLE_PRIMARY_KEY, dbs

    ' GetTableSortFields falls back to every field here, which is deliberately not
    ' something the merge key is allowed to do.
    CreateTestTable dbs, TEST_TABLE_SORTFIELDS, _
        "CREATE TABLE [" & TEST_TABLE_SORTFIELDS & "] (Alpha TEXT(10), Beta LONG)"
    TestAssert GetTableMergeKey(dbs.TableDefs(TEST_TABLE_SORTFIELDS)).Count = 0, _
        "table with no unique index has no merge key"
    TestAssert GetTableSortFields(dbs.TableDefs(TEST_TABLE_SORTFIELDS)).Count = 2, _
        "sort fields still fall back to all fields"
    DropTestTable TEST_TABLE_SORTFIELDS, dbs

End Sub


Public Sub TestGetTableMergeStrategy()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strReason As String

    Set dbs = CurrentDb

    CreateTestTable dbs, TEST_TABLE_PLAIN, _
        "CREATE TABLE [" & TEST_TABLE_PLAIN & "] (ID LONG PRIMARY KEY, Alpha TEXT(10))"
    TestAssert GetTableMergeStrategy(dbs.TableDefs(TEST_TABLE_PLAIN), strReason) = etmsReconcile, _
        "keyed table of simple types is reconciled row by row"
    TestAssert Len(strReason) = 0, "no reason given when mergeable"
    DropTestTable TEST_TABLE_PLAIN, dbs

    ' Without a key there is nothing to pair rows on, so the rows are replaced instead.
    CreateTestTable dbs, TEST_TABLE_SORTFIELDS, _
        "CREATE TABLE [" & TEST_TABLE_SORTFIELDS & "] (Alpha TEXT(10), Beta LONG)"
    TestAssert GetTableMergeStrategy(dbs.TableDefs(TEST_TABLE_SORTFIELDS), strReason) = etmsReload, _
        "table with no unique index is reloaded"
    TestAssert Len(strReason) = 0, "no reason given when reloadable"
    DropTestTable TEST_TABLE_SORTFIELDS, dbs

    CreateTestTable dbs, TEST_TABLE_BINARY, _
        "CREATE TABLE [" & TEST_TABLE_BINARY & "] (ID LONG PRIMARY KEY, Blob LONGBINARY)"
    TestAssert GetTableMergeStrategy(dbs.TableDefs(TEST_TABLE_BINARY), strReason) = etmsNone, _
        "binary field blocks any merge"
    TestAssert InStr(1, strReason, "Blob", vbTextCompare) > 0, "reason names the binary field"
    DropTestTable TEST_TABLE_BINARY, dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestGetTableMergeStrategy_KeylessWithDependent
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : A keyless table can only be reloaded, and a reload deletes every row. When
'           : another table references it the delete would fail and roll the table back,
'           : so the merge is refused up front with a message naming the blocking table.
'---------------------------------------------------------------------------------------
'
Public Sub TestGetTableMergeStrategy_KeylessWithDependent()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strReason As String

    Set dbs = CurrentDb
    DropTestTable TEST_TABLE_MERGE_CHILD, dbs
    ' A unique index on a nullable column: enough for a relationship to reference, but not
    ' a merge key, since a Null could match several rows.
    CreateTestTable dbs, TEST_TABLE_MERGE_PARENT, _
        "CREATE TABLE [" & TEST_TABLE_MERGE_PARENT & "] (ID LONG, Label TEXT(50))"
    dbs.Execute "CREATE UNIQUE INDEX [uq_parent] ON [" & TEST_TABLE_MERGE_PARENT & "] (ID)", _
        dbFailOnError
    dbs.TableDefs.Refresh
    TestAssert GetTableMergeKey(dbs.TableDefs(TEST_TABLE_MERGE_PARENT)).Count = 0, _
        "a unique index on a nullable column is not a merge key"
    TestAssert GetTableMergeStrategy(dbs.TableDefs(TEST_TABLE_MERGE_PARENT), strReason) = etmsReload, _
        "keyless table with nothing referencing it is reloaded"

    CreateTestTable dbs, TEST_TABLE_MERGE_CHILD, _
        "CREATE TABLE [" & TEST_TABLE_MERGE_CHILD & "] (CID LONG PRIMARY KEY, ParentID LONG" & _
        " REFERENCES [" & TEST_TABLE_MERGE_PARENT & "] (ID))"
    TestAssert GetFirstDependentTable(TEST_TABLE_MERGE_PARENT) = TEST_TABLE_MERGE_CHILD, _
        "the referencing table is reported as a dependent"
    TestAssert Len(GetFirstDependentTable(TEST_TABLE_MERGE_CHILD)) = 0, _
        "the referencing table itself has no dependents"
    TestAssert GetTableMergeStrategy(dbs.TableDefs(TEST_TABLE_MERGE_PARENT), strReason) = etmsNone, _
        "a reload is refused while another table references the rows"
    TestAssert InStr(1, strReason, TEST_TABLE_MERGE_CHILD, vbTextCompare) > 0, _
        "reason names the table that blocks the reload"

    DropTestTable TEST_TABLE_MERGE_CHILD, dbs
    DropTestTable TEST_TABLE_MERGE_PARENT, dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableDataMerge_ReconcilesRowsFromTdf
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : One merge has to add a missing row, update a changed row, and remove a row
'           : the source no longer has, and a second merge of the same file has to be a
'           : no-op.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDataMerge_ReconcilesRowsFromTdf()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strFile As String
    Dim strExpected As String

    Set dbs = CurrentDb
    CreateTestTable dbs, TEST_TABLE_MERGE, _
        "CREATE TABLE [" & TEST_TABLE_MERGE & "] (ID LONG PRIMARY KEY, Label TEXT(50))"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE & "] (ID, Label) VALUES (1, 'A')", dbFailOnError
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE & "] (ID, Label) VALUES (2, 'B')", dbFailOnError
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE & "] (ID, Label) VALUES (3, 'C')", dbFailOnError

    ' The exported file becomes the state the merge has to restore.
    strFile = GetTestSourceFile(TEST_TABLE_MERGE, "txt")
    ExportTestTableData TEST_TABLE_MERGE, etdTabDelimited, strFile
    strExpected = GetRowSummary("SELECT ID, Label FROM [" & TEST_TABLE_MERGE & "] ORDER BY ID")

    ' Diverge in all three directions at once.
    dbs.Execute "UPDATE [" & TEST_TABLE_MERGE & "] SET Label = 'changed' WHERE ID = 1", dbFailOnError
    dbs.Execute "DELETE FROM [" & TEST_TABLE_MERGE & "] WHERE ID = 2", dbFailOnError
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE & "] (ID, Label) VALUES (4, 'D')", dbFailOnError

    MergeTestTableData etdTabDelimited, strFile
    TestAssert GetRowSummary("SELECT ID, Label FROM [" & TEST_TABLE_MERGE & "] ORDER BY ID") = strExpected, _
        "row added, row updated, and row removed to match the source file"

    MergeTestTableData etdTabDelimited, strFile
    TestAssert GetRowSummary("SELECT ID, Label FROM [" & TEST_TABLE_MERGE & "] ORDER BY ID") = strExpected, _
        "merging the same file again changes nothing"

    AssertNoStagingTables

    DeleteTestSourceFile strFile
    DropTestTable TEST_TABLE_MERGE, dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableDataMerge_EmptyXmlSourceEmptiesTable
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : A source file with no records means the table should hold none. The load path
'           : skips ImportXML in this case, so the empty staging table has to be reached
'           : anyway -- and the merge has to finish promptly, since ImportXML takes about
'           : 95 seconds on a row-less document.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDataMerge_EmptyXmlSourceEmptiesTable()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strFile As String
    Dim strSql As String
    Dim sngStart As Single

    Set dbs = CurrentDb
    CreateTestTable dbs, TEST_TABLE_MERGE, _
        "CREATE TABLE [" & TEST_TABLE_MERGE & "] (ID LONG PRIMARY KEY, Label TEXT(50))"
    strSql = "SELECT ID, Label FROM [" & TEST_TABLE_MERGE & "] ORDER BY ID"

    ' Export while empty, then add rows the source file does not have.
    strFile = GetTestSourceFile(TEST_TABLE_MERGE, "xml")
    ExportTestTableData TEST_TABLE_MERGE, etdXML, strFile
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE & "] (ID, Label) VALUES (1, 'A')", dbFailOnError
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE & "] (ID, Label) VALUES (2, 'B')", dbFailOnError

    sngStart = Timer
    MergeTestTableData etdXML, strFile
    TestAssert Len(GetRowSummary(strSql)) = 0, "table emptied to match a source file with no records"
    TestAssert Timer - sngStart < 10, "row-less source merged without the ImportXML stall"

    AssertNoStagingTables

    DeleteTestSourceFile strFile
    DropTestTable TEST_TABLE_MERGE, dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableDataMerge_ReconcilesRowsFromXml
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : The XML rows have to be relabeled to load into the staging table. This table
'           : has a field named after the table itself, which is the case that a textual
'           : tag replacement would corrupt.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDataMerge_ReconcilesRowsFromXml()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strFile As String
    Dim strSql As String
    Dim strExpected As String

    Set dbs = CurrentDb
    CreateTestTable dbs, TEST_TABLE_MERGE_XML, _
        "CREATE TABLE [" & TEST_TABLE_MERGE_XML & "] (ID LONG PRIMARY KEY, [" & _
        TEST_TABLE_MERGE_XML & "] TEXT(50))"
    strSql = "SELECT ID, [" & TEST_TABLE_MERGE_XML & "] FROM [" & TEST_TABLE_MERGE_XML & "] ORDER BY ID"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE_XML & "] (ID, [" & TEST_TABLE_MERGE_XML & _
        "]) VALUES (1, 'A')", dbFailOnError
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE_XML & "] (ID, [" & TEST_TABLE_MERGE_XML & _
        "]) VALUES (2, 'B')", dbFailOnError

    strFile = GetTestSourceFile(TEST_TABLE_MERGE_XML, "xml")
    ExportTestTableData TEST_TABLE_MERGE_XML, etdXML, strFile
    strExpected = GetRowSummary(strSql)

    dbs.Execute "UPDATE [" & TEST_TABLE_MERGE_XML & "] SET [" & TEST_TABLE_MERGE_XML & _
        "] = 'changed' WHERE ID = 1", dbFailOnError
    dbs.Execute "DELETE FROM [" & TEST_TABLE_MERGE_XML & "] WHERE ID = 2", dbFailOnError

    MergeTestTableData etdXML, strFile
    TestAssert GetRowSummary(strSql) = strExpected, _
        "XML rows merged even though a field shares the table name"

    AssertNoStagingTables

    DeleteTestSourceFile strFile
    DropTestTable TEST_TABLE_MERGE_XML, dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableDataMerge_CompositeKey
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : The engine only accepts an UPDATE across a join when the joined side is
'           : provably unique, which for a composite key depends on the multi-column
'           : unique index the staging table is given.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDataMerge_CompositeKey()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strFile As String
    Dim strSql As String
    Dim strExpected As String

    Set dbs = CurrentDb
    CreateTestTable dbs, TEST_TABLE_MERGE_COMPOSITE, _
        "CREATE TABLE [" & TEST_TABLE_MERGE_COMPOSITE & "] (ID LONG, Kind LONG, Label TEXT(50)," & _
        " CONSTRAINT PK PRIMARY KEY (ID, Kind))"
    strSql = "SELECT ID, Kind, Label FROM [" & TEST_TABLE_MERGE_COMPOSITE & "] ORDER BY ID, Kind"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE_COMPOSITE & "] (ID, Kind, Label) VALUES (1, 1, 'A')", dbFailOnError
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE_COMPOSITE & "] (ID, Kind, Label) VALUES (1, 2, 'B')", dbFailOnError
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE_COMPOSITE & "] (ID, Kind, Label) VALUES (2, 1, 'C')", dbFailOnError

    strFile = GetTestSourceFile(TEST_TABLE_MERGE_COMPOSITE, "txt")
    ExportTestTableData TEST_TABLE_MERGE_COMPOSITE, etdTabDelimited, strFile
    strExpected = GetRowSummary(strSql)

    dbs.Execute "UPDATE [" & TEST_TABLE_MERGE_COMPOSITE & "] SET Label = 'changed' WHERE ID = 1 AND Kind = 2", dbFailOnError
    dbs.Execute "DELETE FROM [" & TEST_TABLE_MERGE_COMPOSITE & "] WHERE ID = 2", dbFailOnError

    MergeTestTableData etdTabDelimited, strFile
    TestAssert GetRowSummary(strSql) = strExpected, "composite key rows reconciled"

    AssertNoStagingTables

    DeleteTestSourceFile strFile
    DropTestTable TEST_TABLE_MERGE_COMPOSITE, dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableDataMerge_NullsAndLongMemo
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Null is not equal to anything, including another Null, so a change to or
'           : from Null is only detected by the explicit null test in the comparison. Memo
'           : values are checked past the 255 character mark, where a truncating
'           : comparison would silently report no change.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDataMerge_NullsAndLongMemo()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strFile As String
    Dim strSql As String
    Dim strExpected As String
    Dim strLong As String

    Set dbs = CurrentDb
    ' [Note] stays bracketed throughout: NOTE is a type keyword in Access DDL (a synonym
    ' for MEMO), so an unbracketed column of that name is a syntax error. Keeping the name
    ' also gives the reconcile a reserved word to prove it brackets every identifier.
    CreateTestTable dbs, TEST_TABLE_MERGE_NULLS, _
        "CREATE TABLE [" & TEST_TABLE_MERGE_NULLS & "] (ID LONG PRIMARY KEY, [Note] MEMO, Amount DOUBLE)"
    strSql = "SELECT ID, [Note], Amount FROM [" & TEST_TABLE_MERGE_NULLS & "] ORDER BY ID"
    strLong = String$(299, "y") & "z"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE_NULLS & "] (ID, [Note], Amount) VALUES (1, '" & _
        strLong & "', 1.5)", dbFailOnError
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE_NULLS & "] (ID, [Note], Amount) VALUES (2, Null, Null)", dbFailOnError

    strFile = GetTestSourceFile(TEST_TABLE_MERGE_NULLS, "xml")
    ExportTestTableData TEST_TABLE_MERGE_NULLS, etdXML, strFile
    strExpected = GetRowSummary(strSql)

    ' Differ only in the last character of the memo, and in both directions across Null.
    dbs.Execute "UPDATE [" & TEST_TABLE_MERGE_NULLS & "] SET [Note] = '" & String$(300, "y") & _
        "', Amount = Null WHERE ID = 1", dbFailOnError
    dbs.Execute "UPDATE [" & TEST_TABLE_MERGE_NULLS & "] SET [Note] = 'now set', Amount = 9 WHERE ID = 2", dbFailOnError

    MergeTestTableData etdXML, strFile
    TestAssert GetRowSummary(strSql) = strExpected, _
        "long memo difference and both Null directions reconciled"

    AssertNoStagingTables

    DeleteTestSourceFile strFile
    DropTestTable TEST_TABLE_MERGE_NULLS, dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableDataMerge_WideTableUpdatesInGroups
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Past a certain width the database engine refuses to compile an update that
'           : assigns and compares every field at once, answering "Query is too complex"
'           : however few rows are involved. The reconcile has to split the fields into
'           : groups and still land on exactly the rows the source file describes.
'           :
'           : Row 1 differs in the first and last field, which fall in different groups, so
'           : this also covers a row that more than one group has to change.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDataMerge_WideTableUpdatesInGroups()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strFile As String
    Dim strSql As String
    Dim strTable As String
    Dim strExpected As String

    Set dbs = CurrentDb
    strTable = "[" & TEST_TABLE_MERGE_WIDE & "]"
    CreateTestTable dbs, TEST_TABLE_MERGE_WIDE, BuildWideTableSql
    strSql = "SELECT * FROM " & strTable & " ORDER BY ID"
    InsertWideRow dbs, 1
    InsertWideRow dbs, 2
    InsertWideRow dbs, 3

    ' The exported file becomes the state the merge has to restore.
    strFile = GetTestSourceFile(TEST_TABLE_MERGE_WIDE, "xml")
    ExportTestTableData TEST_TABLE_MERGE_WIDE, etdXML, strFile
    strExpected = GetRowSummary(strSql)

    ' Diverge in every direction the reconcile has to handle, spread across field groups.
    dbs.Execute "UPDATE " & strTable & " SET [" & WideFieldName(1) & "] = 'changed', [" & _
        WideFieldName(WIDE_FIELD_COUNT) & "] = 'changed' WHERE ID = 1", dbFailOnError
    dbs.Execute "UPDATE " & strTable & " SET [" & WideFieldName(40) & "] = Null, [" & _
        WideFieldName(2) & "] = 'now set' WHERE ID = 2", dbFailOnError
    dbs.Execute "DELETE FROM " & strTable & " WHERE ID = 3", dbFailOnError
    InsertWideRow dbs, 4

    MergeTestTableData etdXML, strFile
    TestAssert GetRowSummary(strSql) = strExpected, _
        "wide table reconciled even though the update had to be split into field groups"

    MergeTestTableData etdXML, strFile
    TestAssert GetRowSummary(strSql) = strExpected, _
        "merging the same file again changes nothing"

    AssertNoStagingTables

    DeleteTestSourceFile strFile
    DropTestTable TEST_TABLE_MERGE_WIDE, dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableDataMerge_MissingSourceFileKeepsRows
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : A table dropped from the export options, or a deleted file, arrives here as
'           : a missing file. Emptying the table is never the right response.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDataMerge_MissingSourceFileKeepsRows()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strFile As String
    Dim strSql As String
    Dim strExpected As String

    Set dbs = CurrentDb
    CreateTestTable dbs, TEST_TABLE_MERGE, _
        "CREATE TABLE [" & TEST_TABLE_MERGE & "] (ID LONG PRIMARY KEY, Label TEXT(50))"
    strSql = "SELECT ID, Label FROM [" & TEST_TABLE_MERGE & "] ORDER BY ID"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE & "] (ID, Label) VALUES (1, 'A')", dbFailOnError
    strExpected = GetRowSummary(strSql)

    ' Never written, standing in for a file that was deleted or is no longer exported.
    strFile = GetTestSourceFile(TEST_TABLE_MERGE, "txt")
    TestAssert Not FSO.FileExists(strFile), "source file does not exist"

    MergeTestTableData etdTabDelimited, strFile
    TestAssert GetRowSummary(strSql) = strExpected, "rows left alone when the source file is gone"

    DeleteTestSourceFile strFile
    DropTestTable TEST_TABLE_MERGE, dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableDataMerge_ReloadsTableWithoutMergeKey
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Without a key a source row cannot be paired with a table row, so every row is
'           : replaced. The end state still has to match the source file exactly, including
'           : dropping a row the file does not have and keeping duplicates the file does.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDataMerge_ReloadsTableWithoutMergeKey()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strFile As String
    Dim strSql As String
    Dim strExpected As String

    Set dbs = CurrentDb
    CreateTestTable dbs, TEST_TABLE_SORTFIELDS, _
        "CREATE TABLE [" & TEST_TABLE_SORTFIELDS & "] (Alpha TEXT(10), Beta LONG)"
    strSql = "SELECT Alpha, Beta FROM [" & TEST_TABLE_SORTFIELDS & "] ORDER BY Alpha, Beta"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_SORTFIELDS & "] (Alpha, Beta) VALUES ('a', 1)", dbFailOnError
    ' Duplicate rows are legal here and have to survive the round trip.
    dbs.Execute "INSERT INTO [" & TEST_TABLE_SORTFIELDS & "] (Alpha, Beta) VALUES ('a', 1)", dbFailOnError
    dbs.Execute "INSERT INTO [" & TEST_TABLE_SORTFIELDS & "] (Alpha, Beta) VALUES (Null, Null)", dbFailOnError

    strFile = GetTestSourceFile(TEST_TABLE_SORTFIELDS, "txt")
    ExportTestTableData TEST_TABLE_SORTFIELDS, etdTabDelimited, strFile
    strExpected = GetRowSummary(strSql)

    ' Diverge in both directions: a row the file does not have, and a row it does.
    dbs.Execute "INSERT INTO [" & TEST_TABLE_SORTFIELDS & "] (Alpha, Beta) VALUES ('b', 2)", dbFailOnError
    dbs.Execute "DELETE FROM [" & TEST_TABLE_SORTFIELDS & "] WHERE Alpha Is Null", dbFailOnError

    MergeTestTableData etdTabDelimited, strFile
    TestAssert GetRowSummary(strSql) = strExpected, "keyless table reloaded to match the source"

    ' A second merge of the same file has to leave the same rows behind.
    MergeTestTableData etdTabDelimited, strFile
    TestAssert GetRowSummary(strSql) = strExpected, "reload is idempotent"

    AssertNoStagingTables

    DeleteTestSourceFile strFile
    DropTestTable TEST_TABLE_SORTFIELDS, dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableDataMerge_SkipsKeylessTableWithDependent
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Reloading a keyless table deletes every row, which fails when another table
'           : references them. The table is left untouched rather than rolled back.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDataMerge_SkipsKeylessTableWithDependent()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strFile As String
    Dim strSql As String
    Dim strDiverged As String

    Set dbs = CurrentDb
    DropTestTable TEST_TABLE_MERGE_CHILD, dbs
    CreateTestTable dbs, TEST_TABLE_MERGE_PARENT, _
        "CREATE TABLE [" & TEST_TABLE_MERGE_PARENT & "] (ID LONG, Label TEXT(50))"
    dbs.Execute "CREATE UNIQUE INDEX [uq_parent] ON [" & TEST_TABLE_MERGE_PARENT & "] (ID)", _
        dbFailOnError
    dbs.TableDefs.Refresh
    strSql = "SELECT ID, Label FROM [" & TEST_TABLE_MERGE_PARENT & "] ORDER BY ID"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE_PARENT & "] (ID, Label) VALUES (1, 'A')", dbFailOnError

    strFile = GetTestSourceFile(TEST_TABLE_MERGE_PARENT, "txt")
    ExportTestTableData TEST_TABLE_MERGE_PARENT, etdTabDelimited, strFile

    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE_PARENT & "] (ID, Label) VALUES (2, 'B')", dbFailOnError
    CreateTestTable dbs, TEST_TABLE_MERGE_CHILD, _
        "CREATE TABLE [" & TEST_TABLE_MERGE_CHILD & "] (CID LONG PRIMARY KEY, ParentID LONG" & _
        " REFERENCES [" & TEST_TABLE_MERGE_PARENT & "] (ID))"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE_CHILD & "] (CID, ParentID) VALUES (1, 2)", dbFailOnError
    strDiverged = GetRowSummary(strSql)

    MergeTestTableData etdTabDelimited, strFile, True
    TestAssert GetRowSummary(strSql) = strDiverged, _
        "keyless table with a dependent is left untouched"

    AssertNoStagingTables

    DeleteTestSourceFile strFile
    DropTestTable TEST_TABLE_MERGE_CHILD, dbs
    DropTestTable TEST_TABLE_MERGE_PARENT, dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTableDataMerge_RollsBackWhenDeleteBlocked
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : A row the source no longer has may still be referenced by a child record.
'           : The reconcile runs in one transaction so the table is left exactly as it
'           : was rather than half merged.
'---------------------------------------------------------------------------------------
'
Public Sub TestTableDataMerge_RollsBackWhenDeleteBlocked()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim strFile As String
    Dim strSql As String
    Dim strDiverged As String

    Set dbs = CurrentDb
    DropTestTable TEST_TABLE_MERGE_CHILD, dbs
    CreateTestTable dbs, TEST_TABLE_MERGE_PARENT, _
        "CREATE TABLE [" & TEST_TABLE_MERGE_PARENT & "] (ID LONG PRIMARY KEY, Label TEXT(50))"
    strSql = "SELECT ID, Label FROM [" & TEST_TABLE_MERGE_PARENT & "] ORDER BY ID"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE_PARENT & "] (ID, Label) VALUES (1, 'A')", dbFailOnError

    strFile = GetTestSourceFile(TEST_TABLE_MERGE_PARENT, "txt")
    ExportTestTableData TEST_TABLE_MERGE_PARENT, etdTabDelimited, strFile

    ' Add a parent row the source file does not have, and pin it with a child record.
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE_PARENT & "] (ID, Label) VALUES (2, 'B')", dbFailOnError
    CreateTestTable dbs, TEST_TABLE_MERGE_CHILD, _
        "CREATE TABLE [" & TEST_TABLE_MERGE_CHILD & "] (CID LONG PRIMARY KEY, ParentID LONG" & _
        " REFERENCES [" & TEST_TABLE_MERGE_PARENT & "] (ID))"
    dbs.Execute "INSERT INTO [" & TEST_TABLE_MERGE_CHILD & "] (CID, ParentID) VALUES (1, 2)", dbFailOnError
    ' Also give the merge an update to perform, so a partial merge would be visible.
    dbs.Execute "UPDATE [" & TEST_TABLE_MERGE_PARENT & "] SET Label = 'changed' WHERE ID = 1", dbFailOnError
    strDiverged = GetRowSummary(strSql)

    MergeTestTableData etdTabDelimited, strFile, True
    TestAssert GetRowSummary(strSql) = strDiverged, _
        "blocked delete rolls back the whole reconcile, including the update"

    AssertNoStagingTables

    DeleteTestSourceFile strFile
    DropTestTable TEST_TABLE_MERGE_CHILD, dbs
    DropTestTable TEST_TABLE_MERGE_PARENT, dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MergeTestTableData
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Merge one source file through clsDbTableData, which reads the table name
'           : from the file name. The change index is disabled so temp test tables never
'           : reach the project index, and expected warnings are routed to the log rather
'           : than to a dialog that would stall an unattended run.
'---------------------------------------------------------------------------------------
'
Private Sub MergeTestTableData(intFormat As eTableDataExportFormat, strFile As String, _
    Optional blnExpectLogError As Boolean)

    Dim cTable As clsDbTableData
    Dim blnIndexDisabled As Boolean
    Dim blnLogActive As Boolean
    Dim intErrorLevel As eErrorLevel

    blnIndexDisabled = VCSIndex.Disabled
    VCSIndex.Disabled = True
    blnLogActive = Log.Active
    intErrorLevel = Operation.ErrorLevel
    If blnExpectLogError Then Log.Active = True

    Set cTable = New clsDbTableData
    cTable.Format = intFormat
    cTable.Parent.Merge strFile

    If blnExpectLogError Then
        Log.Active = blnLogActive
        ' The warning was the expected outcome, so it should not color the test run.
        Operation.ErrorLevel = intErrorLevel
    End If
    VCSIndex.Disabled = blnIndexDisabled

End Sub


'---------------------------------------------------------------------------------------
' Function  : GetRowSummary
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Render the rows of a query as one comparable string, distinguishing Null
'           : from an empty value.
'---------------------------------------------------------------------------------------
'
Private Function GetRowSummary(strSql As String) As String

    Dim rst As DAO.Recordset
    Dim fld As DAO.Field
    Dim cData As clsConcat

    Set cData = New clsConcat
    Set rst = CurrentDb.OpenRecordset(strSql, dbOpenSnapshot, dbReadOnly)
    Do While Not rst.EOF
        For Each fld In rst.Fields
            cData.Add CStr(Nz(fld.Value, "<null>")), ":"
        Next fld
        cData.Add "|"
        rst.MoveNext
    Loop
    rst.Close

    GetRowSummary = cData.GetStr

End Function


'---------------------------------------------------------------------------------------
' Function  : GetTestSourceFile
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Return a path in a fresh temporary folder for a table data source file.
'           : Merge reads the table name from the file name, so a random temporary file
'           : name will not do; the folder is what makes the name collision-free.
'---------------------------------------------------------------------------------------
'
Private Function GetTestSourceFile(strTable As String, strExt As String) As String
    GetTestSourceFile = GetTempFolder("VCS") & PathSep & strTable & "." & strExt
End Function


'---------------------------------------------------------------------------------------
' Procedure : DeleteTestSourceFile
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Remove a source file created by GetTestSourceFile, along with its folder.
'---------------------------------------------------------------------------------------
'
Private Sub DeleteTestSourceFile(strFile As String)

    Dim strFolder As String

    strFolder = FSO.GetParentFolderName(strFile)
    DeleteFile strFile

    LogUnhandledErrors
    On Error Resume Next
    If FSO.FolderExists(strFolder) Then FSO.DeleteFolder strFolder, True
    If Err Then Err.Clear
    On Error GoTo 0

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AssertNoStagingTables
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : The staging table is temporary. One left behind would be picked up as a
'           : table definition by the next export.
'---------------------------------------------------------------------------------------
'
Private Sub AssertNoStagingTables()
    TestAssert DCount("*", "MSysObjects", "Name Like 'vcs_tmp_merge_data*'") = 0, _
        "no staging table left behind"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CreateTestTable
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Drop and re-create a temp table, refreshing the collections that would
'           : otherwise keep reporting the table as missing.
'---------------------------------------------------------------------------------------
'
Private Sub CreateTestTable(dbs As DAO.Database, strTable As String, strSql As String)

    DropTestTable strTable, dbs
    dbs.Execute strSql, dbFailOnError
    RefreshTableCollections dbs
    TestAssert TableExists(strTable, dbs), "test table created: " & strTable

End Sub


'---------------------------------------------------------------------------------------
' Function  : GetTestTableAccessObject
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Return a CurrentData.AllTables item after CreateTestTable.
'---------------------------------------------------------------------------------------
'
Private Function GetTestTableAccessObject(strTable As String) As AccessObject

    Dim tbl As AccessObject

    For Each tbl In CurrentData.AllTables
        If StrComp(tbl.Name, strTable, vbTextCompare) = 0 Then
            Set GetTestTableAccessObject = tbl
            Exit Function
        End If
    Next tbl

End Function


'---------------------------------------------------------------------------------------
' Procedure : ExportTestTableData
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Export one table to a temp file using clsDbTableData.
'---------------------------------------------------------------------------------------
'
Private Sub ExportTestTableData(strTable As String, intFormat As eTableDataExportFormat, strFile As String)

    Dim cTable As clsDbTableData
    Dim tbl As AccessObject

    Set tbl = GetTestTableAccessObject(strTable)
    If tbl Is Nothing Then
        TestAssert False, "table not found for export: " & strTable
        Exit Sub
    End If

    Set cTable = New clsDbTableData
    cTable.Format = intFormat
    Set cTable.Parent.DbObject = tbl
    cTable.Parent.Export strFile
    ReleaseTableDataSortExport

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DropTestTable
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Drop a temp table if it exists.
'---------------------------------------------------------------------------------------
'
Private Sub DropTestTable(strTable As String, Optional dbs As DAO.Database)

    If dbs Is Nothing Then Set dbs = CurrentDb

    If Not TableExists(strTable, dbs) Then Exit Sub
    dbs.Execute "DROP TABLE [" & strTable & "]", dbFailOnError
    RefreshTableCollections dbs

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RefreshTableCollections
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Pick up a schema change made through DAO. Releasing the shared reference
'           : matters most: clsDbTableData reads the table through SharedDb, and that
'           : cached handle raises error 3265 for any table created after it was opened.
'---------------------------------------------------------------------------------------
'
Private Sub RefreshTableCollections(dbs As DAO.Database)
    dbs.TableDefs.Refresh
    ReleaseDbReferences
End Sub


'---------------------------------------------------------------------------------------
' Function  : BuildWideTableSql
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Return the DDL for the wide merge test table. The text fields are kept short
'           : so their declared sizes stay well inside the 4,000 character record limit.
'---------------------------------------------------------------------------------------
'
Private Function BuildWideTableSql() As String

    Dim cSql As clsConcat
    Dim lngField As Long

    Set cSql = New clsConcat
    cSql.Add "CREATE TABLE [", TEST_TABLE_MERGE_WIDE, "] (ID LONG PRIMARY KEY"
    For lngField = 1 To WIDE_FIELD_COUNT
        cSql.Add ", [", WideFieldName(lngField), "] TEXT(20)"
    Next lngField
    cSql.Add ")"

    BuildWideTableSql = cSql.GetStr

End Function


'---------------------------------------------------------------------------------------
' Function  : WideFieldName
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Return the name of a field in the wide merge test table.
'---------------------------------------------------------------------------------------
'
Private Function WideFieldName(lngField As Long) As String
    WideFieldName = "F" & Format$(lngField, "000")
End Function


'---------------------------------------------------------------------------------------
' Procedure : InsertWideRow
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Add a row to the wide merge test table, with a value in every field except
'           : the one matching the row number. That Null gives the merge a change out of
'           : Null to detect, which an inequality test on its own would miss.
'---------------------------------------------------------------------------------------
'
Private Sub InsertWideRow(dbs As DAO.Database, lngId As Long)

    Dim cSql As clsConcat
    Dim cValues As clsConcat
    Dim lngField As Long

    Set cSql = New clsConcat
    Set cValues = New clsConcat
    cSql.Add "INSERT INTO [", TEST_TABLE_MERGE_WIDE, "] (ID"
    cValues.Add CStr(lngId)

    For lngField = 1 To WIDE_FIELD_COUNT
        cSql.Add ", [", WideFieldName(lngField), "]"
        If lngField = lngId Then
            cValues.Add ", Null"
        Else
            cValues.Add ", '", WideFieldName(lngField), "-", CStr(lngId), "'"
        End If
    Next lngField

    cSql.Add ") VALUES (", cValues.GetStr, ")"
    dbs.Execute cSql.GetStr, dbFailOnError

End Sub


'---------------------------------------------------------------------------------------
' Function  : BuildSingleTextXmlSortKey
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Build a one-field ComposeXmlSortKey for text sort regression tests.
'---------------------------------------------------------------------------------------
'
Private Function BuildSingleTextXmlSortKey(strValue As String, lngOrdinal As Long) As String

    Dim colParts As New Collection

    colParts.Add NormalizeXmlSortValue(strValue, dbText)
    BuildSingleTextXmlSortKey = ComposeXmlSortKey(colParts, lngOrdinal)

End Function
