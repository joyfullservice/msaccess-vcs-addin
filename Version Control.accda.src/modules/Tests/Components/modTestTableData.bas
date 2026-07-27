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
Private Const TEST_TABLE_SORTFIELDS As String = "vcs_test_sortfields"
Private Const TEST_TABLE_PRIMARY_KEY As String = "vcs_test_sortfields_pk"


Public Sub TestEscapeXmlName()
    TestAssert EscapeXmlName("NotReq'd") = "NotReq_x0027_d", "apostrophe"
    TestAssert EscapeXmlName("Please" & Chr$(34) & "d" & Chr$(34) & "don" & Chr$(39) & "t" & Chr$(34) & "use") = _
        "Please_x0022_d_x0022_don_x0027_t_x0022_use", "quotes and apostrophes"
    TestAssert EscapeXmlName("ID") = "ID", "simple name unchanged"
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
