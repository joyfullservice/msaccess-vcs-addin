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
Private Const TEST_TABLE_OUTSIDE_EXPORT As String = "vcs_test_td_outside_export"
Private Const TEST_TABLE_DECIMAL_SQL As String = "vcs_test_decimal_sql"
Private Const TEST_TABLE_DECIMAL_DAO As String = "vcs_test_decimal_dao"
Private Const TEST_TABLE_HIDDEN_SIDECAR As String = "vcs_test_hidden_sidecar"
Private Const TEST_TABLE_SYSTEM_PREFIX As String = "MSysVcsTestUserTable"


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
    Dim lngImportError As Long
    Dim strImportError As String

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
    LogUnhandledErrors
    On Error Resume Next
    Application.ImportXML strTemp, acStructureOnly
    lngImportError = Err.Number
    strImportError = Err.Description
    Err.Clear
    On Error GoTo 0
    ReleaseDbReferences

    TestAssert lngImportError = 0 Or lngImportError = 31550, _
        "ImportXML returned an unexpected error " & CStr(lngImportError) & ": " & _
        strImportError
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
' Procedure : TestDaoImportFromOutsideExportFolder
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Pin that the DAO fast path still verifies (and keeps) a table whose source
'           : XML sits outside Options.GetExportFolder. StoredDefinitionMatchesSource used
'           : to treat that case as an automatic mismatch by refusing to write a temp
'           : export when the export-folder prefix rewrite was a no-op, which made every
'           : round-trip fixture under Testing\Fixtures\scratch\ fall through to
'           : Application.ImportXML even when the DAO build was correct.
'---------------------------------------------------------------------------------------
'
Public Sub TestDaoImportFromOutsideExportFolder()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim strXml As String
    Dim cTable As clsDbTableDef
    Dim cComponent As IDbComponent
    Dim lngPriorFormat As Long
    Dim blnPriorOverride As Boolean

    DropTestTable TEST_TABLE_OUTSIDE_EXPORT

    Set dbs = CurrentDb
    Set tdf = dbs.CreateTableDef(TEST_TABLE_OUTSIDE_EXPORT)
    Set fld = tdf.CreateField("ID", dbLong)
    fld.Attributes = fld.Attributes Or dbAutoIncrField
    tdf.Fields.Append fld
    dbs.TableDefs.Append tdf
    RefreshTableCollections dbs

    ' Export through the same path the verification step uses, under the format that
    ' makes DAO and ImportXML output comparable, so a successful re-import is able to
    ' hash-match the source file.
    lngPriorFormat = Options.ExportFormatVersion
    Options.ExportFormatVersion = EFV_5_1_0

    ' Import derives the object name from the file basename, so the source file
    ' must be named for the table. Place it under the system temp directory --
    ' never under the project's export folder -- which is the case under test.
    strXml = GetTempFile
    DeleteFile strXml
    strXml = FSO.GetParentFolderName(strXml) & PathSep & TEST_TABLE_OUTSIDE_EXPORT & ".xml"
    If FSO.FileExists(strXml) Then DeleteFile strXml
    Application.ExportXML acExportTable, TEST_TABLE_OUTSIDE_EXPORT, , strXml, , , , _
        acExportAllTableAndFieldProperties
    With New clsSourceParser
        .LoadSourceFile strXml, edbTableDef
        DeleteFile strXml
        WriteFile .Sanitize(ectXML), strXml
    End With

    DropTestTable TEST_TABLE_OUTSIDE_EXPORT

    TestAssert InStr(1, strXml, Options.GetExportFolder, vbTextCompare) = 0, _
        "fixture path is outside the export folder"

    blnPriorOverride = FastPathTestOverride
    FastPathTestOverride = True

    Set cTable = New clsDbTableDef
    Set cComponent = cTable
    cComponent.Import strXml
    ReleaseDbReferences

    FastPathTestOverride = blnPriorOverride
    Options.ExportFormatVersion = lngPriorFormat

    TestAssert TableExists(TEST_TABLE_OUTSIDE_EXPORT), "import created the table"
    TestAssert Len(GetLastDeclineReason()) = 0, _
        "DAO path kept the table (verification ran outside the export folder)"

    DeleteFile strXml
    DropTestTable TEST_TABLE_OUTSIDE_EXPORT

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestSaveTableSqlDecimalScale
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : Issue 756. SaveTableSqlDef must read DAO Field.Scale, not NumericScale, when
'           : emitting DECIMAL(p,s) in the optional .sql companion file.
'---------------------------------------------------------------------------------------
'
Public Sub TestSaveTableSqlDecimalScale()
    '@Tag("integration")

    Dim cTable As clsDbTableDef
    Dim strFolder As String
    Dim strSqlFile As String
    Dim strSql As String

    DropTestTable TEST_TABLE_DECIMAL_SQL

    CurrentProject.Connection.Execute _
        "CREATE TABLE [" & TEST_TABLE_DECIMAL_SQL & "] (" & _
        "[ID] LONG, [Year4] DECIMAL(4,0), [Amount] DECIMAL(18,4))"
    RefreshTableCollections CurrentDb

    strFolder = GetTempFolder("vcs_dec_sql")

    Set cTable = New clsDbTableDef
    cTable.SaveTableSqlDef TEST_TABLE_DECIMAL_SQL, AddSlash(strFolder)

    strSqlFile = AddSlash(strFolder) & GetSafeFileName(TEST_TABLE_DECIMAL_SQL) & ".sql"
    TestAssert FSO.FileExists(strSqlFile), "sql companion file was written"
    strSql = ReadFile(strSqlFile)

    TestAssert InStr(1, strSql, "DECIMAL(4,0)", vbTextCompare) > 0, "Year4 emits DECIMAL(4,0)"
    TestAssert InStr(1, strSql, "DECIMAL(18,4)", vbTextCompare) > 0, "Amount emits DECIMAL(18,4)"
    TestAssert InStr(1, strSql, "[Year4] VARCHAR", vbTextCompare) = 0, "Year4 is not VARCHAR"
    TestAssert InStr(1, strSql, "[Amount] VARCHAR", vbTextCompare) = 0, "Amount is not VARCHAR"

    If FSO.FolderExists(strFolder) Then FSO.DeleteFolder strFolder, True
    DropTestTable TEST_TABLE_DECIMAL_SQL

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDaoDecimalPrecisionAndScale
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : The DAO table builder must materialize decimal fields as dbDecimal with the
'           : parsed precision and scale, using ADO ALTER after CreateField.
'---------------------------------------------------------------------------------------
'
Public Sub TestDaoDecimalPrecisionAndScale()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim fld As DAO.Field
    Dim strXml As String
    Dim blnPriorOverride As Boolean

    DropTestTable TEST_TABLE_DECIMAL_DAO

    strXml = GetTempFile
    DeleteFile strXml
    strXml = FSO.GetParentFolderName(strXml) & PathSep & TEST_TABLE_DECIMAL_DAO & ".xml"
    If FSO.FileExists(strXml) Then DeleteFile strXml
    WriteFile DecimalFixtureXml, strXml

    blnPriorOverride = FastPathTestOverride
    FastPathTestOverride = True

    TestAssert TryBuildTableFromDefXml(strXml, TEST_TABLE_DECIMAL_DAO), _
        "DAO build succeeds for decimal fields"
    TestAssert Len(GetLastDeclineReason()) = 0, "no decline reason"

    FastPathTestOverride = blnPriorOverride

    Set dbs = CurrentDb
    TestAssert TableExists(TEST_TABLE_DECIMAL_DAO, dbs), "table was created"

    Set fld = dbs.TableDefs(TEST_TABLE_DECIMAL_DAO).Fields("Year4")
    TestAssert fld.Type = dbDecimal, "Year4 type is dbDecimal"
    TestAssert fld.Precision = 4, "Year4 precision"
    TestAssert fld.Scale = 0, "Year4 scale"

    Set fld = dbs.TableDefs(TEST_TABLE_DECIMAL_DAO).Fields("Amount")
    TestAssert fld.Type = dbDecimal, "Amount type is dbDecimal"
    TestAssert fld.Precision = 18, "Amount precision"
    TestAssert fld.Scale = 4, "Amount scale"

    DeleteFile strXml
    DropTestTable TEST_TABLE_DECIMAL_DAO

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


Private Function DecimalFixtureXml() As String

    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "<?xml version=""1.0""?>"
        .Add "<xsd:schema xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:od=""urn:schemas-microsoft-com:officedata"">"
        .Add "  <xsd:element name=""dataroot"">"
        .Add "    <xsd:complexType><xsd:sequence>"
        .Add "      <xsd:element ref=""" & TEST_TABLE_DECIMAL_DAO & """ minOccurs=""0"" maxOccurs=""unbounded""/>"
        .Add "    </xsd:sequence>"
        .Add "    <xsd:attribute name=""generated"" type=""xsd:dateTime""/>"
        .Add "    </xsd:complexType></xsd:element>"
        .Add "  <xsd:element name=""" & TEST_TABLE_DECIMAL_DAO & """>"
        .Add "    <xsd:annotation><xsd:appinfo>"
        .Add "      <od:index index-name=""PrimaryKey"" index-key=""ID "" primary=""yes"" unique=""yes"" clustered=""no"" order=""asc""/>"
        .Add "    </xsd:appinfo></xsd:annotation>"
        .Add "    <xsd:complexType><xsd:sequence>"
        .Add "      <xsd:element name=""ID"" minOccurs=""1"" od:jetType=""autonumber"" od:sqlSType=""int"" od:autoUnique=""yes"" od:nonNullable=""yes"" type=""xsd:int""/>"
        .Add "      <xsd:element name=""Year4"" minOccurs=""0"" od:jetType=""decimal"">"
        .Add "        <xsd:simpleType><xsd:restriction base=""xsd:decimal"">"
        .Add "          <xsd:totalDigits value=""4""/>"
        .Add "          <xsd:fractionDigits value=""0""/>"
        .Add "        </xsd:restriction></xsd:simpleType></xsd:element>"
        .Add "      <xsd:element name=""Amount"" minOccurs=""0"" od:jetType=""decimal"">"
        .Add "        <xsd:simpleType><xsd:restriction base=""xsd:decimal"">"
        .Add "          <xsd:totalDigits value=""18""/>"
        .Add "          <xsd:fractionDigits value=""4""/>"
        .Add "        </xsd:restriction></xsd:simpleType></xsd:element>"
        .Add "    </xsd:sequence></xsd:complexType></xsd:element>"
        .Add "</xsd:schema>"
        DecimalFixtureXml = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : TestHiddenLocalTableKeepsMetadataSidecar
' Author    : Ricardo Hernandez
' Date      : 9/5/2026
' Purpose   : A hidden local table must keep its companion metadata file after a normal
'           : export. The alternate-format name for a local table is <name>.json -- the
'           : same name the metadata sidecar uses -- so RemoveAlternateFormatSourceFile
'           : deletes the file ExportObjectMetadata wrote a few lines earlier, and the
'           : hidden attribute never reaches source. A Build from Source then recreates
'           : the table visible.
'           :
'           : The export folder is redirected to a temp folder so the fixture never
'           : writes into the project's own source tree, and restored before the
'           : assertions so a failure cannot leave the option pointing elsewhere.
'---------------------------------------------------------------------------------------
'
Public Sub TestHiddenLocalTableKeepsMetadataSidecar()
    '@Tag("integration")

    Dim cTable As clsDbTableDef
    Dim cComponent As IDbComponent
    Dim dFile As Dictionary
    Dim dItems As Dictionary
    Dim strFolder As String
    Dim strPriorFolder As String
    Dim lngPriorFormat As Long
    Dim strBase As String
    Dim strXmlFile As String
    Dim strJsonFile As String
    Dim blnWasHidden As Boolean
    Dim blnJsonExists As Boolean
    Dim blnHasHidden As Boolean

    DropTestTable TEST_TABLE_HIDDEN_SIDECAR

    CurrentProject.Connection.Execute _
        "CREATE TABLE [" & TEST_TABLE_HIDDEN_SIDECAR & "] ([ID] LONG)"
    RefreshTableCollections CurrentDb
    Application.SetHiddenAttribute acTable, TEST_TABLE_HIDDEN_SIDECAR, True
    blnWasHidden = Application.GetHiddenAttribute(acTable, TEST_TABLE_HIDDEN_SIDECAR)

    strFolder = GetTempFolder("vcs_hidden_sidecar")
    strPriorFolder = Options.ExportFolder
    lngPriorFormat = Options.ExportFormatVersion
    Options.ExportFolder = AddSlash(strFolder)
    Options.ExportFormatVersion = EFV_5_0_0

    ' Export through the normal (non-alternate) path -- the one a real export takes for
    ' any table the change scan did not already export to the temp folder.
    Set cTable = New clsDbTableDef
    Set cComponent = cTable
    Set cComponent.DbObject = CurrentData.AllTables(TEST_TABLE_HIDDEN_SIDECAR)
    cComponent.Export

    strBase = Options.GetExportFolder & "tbldefs" & PathSep & GetSafeFileName(TEST_TABLE_HIDDEN_SIDECAR)
    strXmlFile = strBase & ".xml"
    strJsonFile = strBase & ".json"
    blnJsonExists = FSO.FileExists(strJsonFile)
    If blnJsonExists Then
        Set dFile = ReadJsonFile(strJsonFile)
        If Not dFile Is Nothing Then
            If dFile.Exists("Items") Then
                Set dItems = dFile("Items")
                If dItems.Exists("Hidden") Then blnHasHidden = dItems("Hidden")
            End If
        End If
    End If

    ' Restore before asserting so a failure cannot leave the options redirected.
    Options.ExportFolder = strPriorFolder
    Options.ExportFormatVersion = lngPriorFormat

    TestAssert blnWasHidden, "fixture table is hidden before the export"
    TestAssert FSO.FileExists(strXmlFile), "local table schema was exported"
    TestAssert blnJsonExists, "metadata sidecar survives the export of a hidden local table"
    TestAssert blnHasHidden, "sidecar records the hidden attribute"

    If FSO.FolderExists(strFolder) Then FSO.DeleteFolder strFolder, True
    DropTestTable TEST_TABLE_HIDDEN_SIDECAR

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestUserTableWithSystemPrefixIsExported
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 9/1/2026
' Purpose   : A user table whose name begins with MSys must be enumerated for export.
'           : The prefix is not reserved, and DAO creates such a table with no system
'           : attribute, which makes it indistinguishable from any other user table.
'           : Checked in both directions on purpose: the engine's own tables must stay
'           : out, so that fixing the omission does not start exporting them instead.
'---------------------------------------------------------------------------------------
'
Public Sub TestUserTableWithSystemPrefixIsExported()
    '@Tag("integration")

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim cCategory As IDbComponent
    Dim cItem As IDbComponent
    Dim dAllTables As Dictionary
    Dim varItem As Variant
    Dim blnFound As Boolean
    Dim blnSystemLeaked As Boolean
    Dim blnNoSystemAttribute As Boolean

    DropTestTable TEST_TABLE_SYSTEM_PREFIX

    Set dbs = CurrentDb
    Set tdf = dbs.CreateTableDef(TEST_TABLE_SYSTEM_PREFIX)
    tdf.Fields.Append tdf.CreateField("ID", dbLong)
    dbs.TableDefs.Append tdf
    RefreshTableCollections dbs

    ' The premise of the fix: the prefix carries no attribute of its own.
    ' Assign the comparison first: a Sub call whose first argument opens with a
    ' parenthesis does not compile in VBA.
    blnNoSystemAttribute = ((dbs.TableDefs(TEST_TABLE_SYSTEM_PREFIX).Attributes And dbSystemObject) = 0)
    TestAssert blnNoSystemAttribute, "a user table named MSys* carries no system attribute"

    ' Access GetAllFromDB through the IDbComponent interface (same pattern as modExport)
    Set cCategory = New clsDbTableDef
    Set dAllTables = cCategory.GetAllFromDB(False)
    For Each varItem In dAllTables.Items
        Set cItem = varItem
        If cItem.Name = TEST_TABLE_SYSTEM_PREFIX Then blnFound = True
        If cItem.Name = "MSysObjects" Then blnSystemLeaked = True
    Next varItem

    TestAssert blnFound, "user table named MSys* should be enumerated for export"
    TestAssert Not blnSystemLeaked, "engine system tables should stay excluded"

    DropTestTable TEST_TABLE_SYSTEM_PREFIX

End Sub


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
