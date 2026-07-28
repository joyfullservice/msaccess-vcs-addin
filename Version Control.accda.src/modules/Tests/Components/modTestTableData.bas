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
