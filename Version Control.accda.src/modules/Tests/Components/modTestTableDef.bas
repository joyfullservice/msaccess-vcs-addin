Attribute VB_Name = "modTestTableDef"
'---------------------------------------------------------------------------------------
' Module    : modTestTableDef
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Unit and integration tests for local table-definition export/import,
'           : including dbBigInt repair after Application.ImportXML.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Components")
'@Tag("unit")


Private Const TEST_TABLE_BIGINT As String = "vcs_test_bigint_repair"


'---------------------------------------------------------------------------------------
' Procedure : TestBigIntFieldXmlDetection
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Access exports a dbBigInt field as an xsd:decimal restriction carrying
'           : `totalDigits value="0"` and no od:jetType attribute. A genuine Decimal
'           : field always carries its real configured precision, so the zero value is
'           : the only available signature.
'---------------------------------------------------------------------------------------
'
Public Sub TestBigIntFieldXmlDetection()
    Dim colNames As Collection
    Dim strXml As String

    strXml = BigIntFixtureXml("BigVal")
    Set colNames = GetBigIntRepairFieldNamesFromTableDefXml(strXml)
    TestAssert colNames.Count = 1, "bigint signature should match one field"
    TestAssert colNames(1) = "BigVal", "matched field name"

    strXml = Replace(strXml, "totalDigits value=""0""", "totalDigits value=""18""")
    Set colNames = GetBigIntRepairFieldNamesFromTableDefXml(strXml)
    TestAssert colNames.Count = 0, "genuine decimal precision should not match"

    strXml = "<xsd:schema xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:od=""urn:schemas-microsoft-com:officedata"">" & _
        "<xsd:element name=""LongVal"" minOccurs=""0"" od:jetType=""longinteger"" od:sqlSType=""int"" type=""xsd:int""/>" & _
        "</xsd:schema>"
    Set colNames = GetBigIntRepairFieldNamesFromTableDefXml(strXml)
    TestAssert colNames.Count = 0, "longinteger field should not match"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestBigIntImportXmlRepair
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : End-to-end check that a dbBigInt field survives the ExportXML/ImportXML
'           : round trip once the ALTER COLUMN repair runs. Each stage re-reads the
'           : table through a fresh Database handle: Application.ImportXML rebuilds the
'           : table outside DAO, so a handle obtained earlier keeps reporting the old
'           : catalog (error 3265) even after TableDefs.Refresh.
'---------------------------------------------------------------------------------------
'
Public Sub TestBigIntImportXmlRepair()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim strTemp As String
    Dim colNames As Collection

    DropTestTable TEST_TABLE_BIGINT

    Set dbs = CurrentDb
    Set tdf = dbs.CreateTableDef(TEST_TABLE_BIGINT)
    Set fld = tdf.CreateField("ID", dbLong)
    tdf.Fields.Append fld
    Set fld = tdf.CreateField("BigVal", dbBigInt)
    tdf.Fields.Append fld
    Set fld = tdf.CreateField("Notes", dbText, 50)
    tdf.Fields.Append fld
    dbs.TableDefs.Append tdf
    RefreshTableCollections dbs

    strTemp = GetTempFile & ".xml"
    Application.ExportXML acExportTable, TEST_TABLE_BIGINT, , strTemp, , , , acExportAllTableAndFieldProperties

    DropTestTable TEST_TABLE_BIGINT
    Application.ImportXML strTemp, acStructureOnly
    ReleaseDbReferences

    Set colNames = GetBigIntRepairFieldNamesFromTableDefXml(ReadFile(strTemp))
    TestAssert colNames.Count = 1, "exported bigint field is detected in source XML"

    If colNames.Count = 1 Then
        Set dbs = CurrentDb
        TestAssert TableExists(TEST_TABLE_BIGINT, dbs), "ImportXML recreated the table"
        TestAssert dbs.TableDefs(TEST_TABLE_BIGINT).Fields(colNames(1)).Type <> dbBigInt, _
            "ImportXML does not preserve dbBigInt on its own"

        dbs.Execute "ALTER TABLE [" & TEST_TABLE_BIGINT & "] ALTER COLUMN [" & _
            colNames(1) & "] BIGINT", dbFailOnError
        RefreshTableCollections dbs

        Set dbs = CurrentDb
        TestAssert dbs.TableDefs(TEST_TABLE_BIGINT).Fields(colNames(1)).Type = dbBigInt, _
            "ALTER COLUMN restores dbBigInt after ImportXML"
        TestAssert dbs.TableDefs(TEST_TABLE_BIGINT).Fields("BigVal").OrdinalPosition = 1, _
            "field ordinal is preserved"
    End If

    DeleteFile strTemp
    DropTestTable TEST_TABLE_BIGINT

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestPropertyNodesSortedUnderNewFormat
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Under EFV_5_1_0 the sanitizer orders od:fieldProperty and od:tableProperty
'           : nodes by name, so a table built through DAO and one built by ImportXML
'           : export identically. Older formats must be left exactly as Access wrote them.
'---------------------------------------------------------------------------------------
'
Public Sub TestPropertyNodesSortedUnderNewFormat()

    Dim strSorted As String
    Dim strLegacy As String
    Dim lngPrior As Long

    lngPrior = Options.ExportFormatVersion

    Options.ExportFormatVersion = EFV_5_1_0
    strSorted = SanitizeFixture(UnsortedPropertyXml)

    Options.ExportFormatVersion = EFV_5_0_0
    strLegacy = SanitizeFixture(UnsortedPropertyXml)

    Options.ExportFormatVersion = lngPrior

    TestAssert InStr(1, strSorted, "AllowZeroLength") < InStr(1, strSorted, "ColumnWidth"), _
        "field properties sorted by name"
    TestAssert InStr(1, strSorted, "ColumnWidth") < InStr(1, strSorted, "Required"), _
        "sort is alphabetical, not merely reversed"
    TestAssert InStr(1, strSorted, "Description") < InStr(1, strSorted, "Orientation"), _
        "table properties sorted by name"

    ' The index shares the appinfo block and must not be dragged into the sort.
    TestAssert InStr(1, strSorted, "od:index") < InStr(1, strSorted, "od:tableProperty"), _
        "index node keeps its position ahead of the properties"

    TestAssert InStr(1, strLegacy, "ColumnWidth") < InStr(1, strLegacy, "AllowZeroLength"), _
        "older export format leaves the original order untouched"

    ' The lookup chain is the one place alphabetical ordering looks dangerous: Access
    ' always emits DisplayControl ahead of BoundColumn and ColumnCount, and issue 691
    ' blamed the reverse for losing ColumnCount and ColumnWidths. Measurement showed the
    ' reverse restores correctly through every import path, so the inversion below is
    ' deliberate. Asserting it keeps a future "fix" from silently reintroducing a rank
    ' table nobody needs. See DECISIONS.md 2026-07-31.
    TestAssert InStr(1, strSorted, "BoundColumn") < InStr(1, strSorted, "DisplayControl"), _
        "lookup chain is alphabetised, inverting Access's emission order on purpose"
    TestAssert InStr(1, strSorted, "ColumnCount") < InStr(1, strSorted, "RowSourceType"), _
        "ColumnCount sorts ahead of RowSourceType"

    ' Reordering must never lose or duplicate a node.
    TestAssert CountOccurrences(strSorted, "<od:fieldProperty ") = _
        CountOccurrences(strLegacy, "<od:fieldProperty "), _
        "sorting preserves the field property count"
    TestAssert CountOccurrences(strSorted, "<od:tableProperty ") = _
        CountOccurrences(strLegacy, "<od:tableProperty "), _
        "sorting preserves the table property count"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestPropertyNodesNotSortedForTableData
' Author    : Adam Waller
' Date      : 7/31/2026
' Purpose   : The sort is gated to table definitions. SanitizeXML is shared with table
'           : data, where a retained schema is rare and the document can run to many
'           : megabytes, so the document scan must not happen there at all.
'---------------------------------------------------------------------------------------
'
Public Sub TestPropertyNodesNotSortedForTableData()

    Dim cParser As clsSourceParser
    Dim strOut As String
    Dim lngPrior As Long

    lngPrior = Options.ExportFormatVersion
    Options.ExportFormatVersion = EFV_5_1_0

    Set cParser = New clsSourceParser
    cParser.LoadString UnsortedPropertyXml, edbTableData
    strOut = cParser.Sanitize(ectXML)

    Options.ExportFormatVersion = lngPrior

    TestAssert InStr(1, strOut, "ColumnWidth") < InStr(1, strOut, "AllowZeroLength"), _
        "table data content is not reordered"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CountOccurrences
' Author    : Adam Waller
' Date      : 7/31/2026
' Purpose   : Count non-overlapping occurrences of a substring.
'---------------------------------------------------------------------------------------
'
Private Function CountOccurrences(strText As String, strFind As String) As Long

    Dim lngPos As Long

    lngPos = InStr(1, strText, strFind, vbBinaryCompare)
    Do While lngPos > 0
        CountOccurrences = CountOccurrences + 1
        lngPos = InStr(lngPos + Len(strFind), strText, strFind, vbBinaryCompare)
    Loop

End Function


'---------------------------------------------------------------------------------------
' Procedure : SanitizeFixture
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Run a table definition XML string through the real export sanitizer.
'---------------------------------------------------------------------------------------
'
Private Function SanitizeFixture(strXml As String) As String

    Dim cParser As clsSourceParser

    Set cParser = New clsSourceParser
    cParser.LoadString strXml, edbTableDef
    SanitizeFixture = cParser.Sanitize(ectXML)

End Function


'---------------------------------------------------------------------------------------
' Procedure : UnsortedPropertyXml
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : A table definition whose properties are deliberately out of name order, in
'           : the arrangement Access produces: Required and AllowZeroLength trailing the
'           : datasheet properties, and the lookup chain led by DisplayControl. Sorting
'           : this by name inverts the lookup chain, which is the case worth pinning.
'---------------------------------------------------------------------------------------
'
' Built through clsConcat rather than a continued expression: VBA allows only 25 line
' continuations in one statement, and this document needs more lines than that.
Private Function UnsortedPropertyXml() As String

    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "<?xml version=""1.0""?>"
        .Add "<xsd:schema xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:od=""urn:schemas-microsoft-com:officedata"">"
        .Add "  <xsd:element name=""Fixture"">"
        .Add "    <xsd:annotation><xsd:appinfo>"
        .Add "      <od:index index-name=""PrimaryKey"" index-key=""F1 "" primary=""yes"" unique=""yes"" clustered=""no"" order=""asc""/>"
        .Add "      <od:tableProperty name=""Orientation"" type=""2"" value=""0""/>"
        .Add "      <od:tableProperty name=""Description"" type=""10"" value=""A table""/>"
        .Add "    </xsd:appinfo></xsd:annotation>"
        .Add "    <xsd:complexType><xsd:sequence>"
        .Add "      <xsd:element name=""F1"" minOccurs=""0"" od:jetType=""text"">"
        .Add "        <xsd:annotation><xsd:appinfo>"
        .Add "          <od:fieldProperty name=""ColumnWidth"" type=""3"" value=""-1""/>"
        .Add "          <od:fieldProperty name=""ColumnOrder"" type=""3"" value=""0""/>"
        .Add "          <od:fieldProperty name=""Required"" type=""1"" value=""0""/>"
        .Add "          <od:fieldProperty name=""AllowZeroLength"" type=""1"" value=""0""/>"
        .Add "          <od:fieldProperty name=""DisplayControl"" type=""3"" value=""111""/>"
        .Add "          <od:fieldProperty name=""RowSourceType"" type=""10"" value=""Value List""/>"
        .Add "          <od:fieldProperty name=""RowSource"" type=""12"" value=""1;One;2;Two""/>"
        .Add "          <od:fieldProperty name=""BoundColumn"" type=""3"" value=""1""/>"
        .Add "          <od:fieldProperty name=""ColumnCount"" type=""3"" value=""2""/>"
        .Add "          <od:fieldProperty name=""ColumnWidths"" type=""10"" value=""0;1440""/>"
        .Add "        </xsd:appinfo></xsd:annotation>"
        .Add "        <xsd:simpleType>"
        .Add "          <xsd:restriction base=""xsd:string"">"
        .Add "            <xsd:maxLength value=""10""/>"
        .Add "          </xsd:restriction>"
        .Add "        </xsd:simpleType>"
        .Add "      </xsd:element>"
        .Add "    </xsd:sequence></xsd:complexType>"
        .Add "  </xsd:element>"
        .Add "</xsd:schema>"
        UnsortedPropertyXml = .GetStr
    End With

End Function


Private Function BigIntFixtureXml(strFieldName As String) As String
    BigIntFixtureXml = _
        "<?xml version=""1.0""?>" & vbCrLf & _
        "<xsd:schema xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:od=""urn:schemas-microsoft-com:officedata"">" & vbCrLf & _
        "  <xsd:element name=""Fixture"">" & vbCrLf & _
        "    <xsd:complexType><xsd:sequence>" & vbCrLf & _
        "      <xsd:element name=""" & strFieldName & """ minOccurs=""0"">" & vbCrLf & _
        "        <xsd:annotation><xsd:appinfo>" & vbCrLf & _
        "          <od:fieldProperty name=""Required"" type=""1"" value=""0""/>" & vbCrLf & _
        "        </xsd:appinfo></xsd:annotation>" & vbCrLf & _
        "        <xsd:simpleType>" & vbCrLf & _
        "          <xsd:restriction base=""xsd:decimal"">" & vbCrLf & _
        "            <xsd:totalDigits value=""0""/>" & vbCrLf & _
        "            <xsd:fractionDigits value=""0""/>" & vbCrLf & _
        "          </xsd:restriction>" & vbCrLf & _
        "        </xsd:simpleType>" & vbCrLf & _
        "      </xsd:element>" & vbCrLf & _
        "    </xsd:sequence></xsd:complexType>" & vbCrLf & _
        "  </xsd:element>" & vbCrLf & _
        "</xsd:schema>"
End Function


'---------------------------------------------------------------------------------------
' Procedure : DropTestTable
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Drop a temp table if it exists. Deliberately avoids `On Error Resume Next`
'           : around an expected "not found" error: the swallowed error stays in Err and
'           : the next LogUnhandledErrors reports it against an unrelated caller.
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
' Date      : 7/30/2026
' Purpose   : Pick up a schema change made through DAO, and release the shared handle so
'           : the next SharedDb caller sees the current catalog.
'---------------------------------------------------------------------------------------
'
Private Sub RefreshTableCollections(dbs As DAO.Database)
    dbs.TableDefs.Refresh
    ReleaseDbReferences
End Sub
