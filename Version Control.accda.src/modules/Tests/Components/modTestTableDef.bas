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
