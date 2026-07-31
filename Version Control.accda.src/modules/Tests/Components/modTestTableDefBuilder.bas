Attribute VB_Name = "modTestTableDefBuilder"
'---------------------------------------------------------------------------------------
' Module    : modTestTableDefBuilder
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Tests for modTableDefBuilder's XML parsing, in isolation from Access.
'           :
'           : These exercise GetTableDefSchemaFromXml, which turns exported table
'           : definition XML into a schema model without touching the database. That
'           : split is what makes the type map testable at all: the round-trip fixtures
'           : in Testing/Fixtures/tabledefs/ prove that real Access output survives a
'           : real create-and-re-export cycle, but they can only cover data types that
'           : happen to appear in the sample database. The cases here cover the rest of
'           : the map from hand-written schema fragments, and every rejection path,
'           : without needing a table of each type to exist somewhere.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests.Components")
'@Tag("unit")


'---------------------------------------------------------------------------------------
' Procedure : TestJetTypeMap
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Every od:jetType we claim to understand maps to the right DAO type.
'---------------------------------------------------------------------------------------
'
Public Sub TestJetTypeMap()

    TestAssert FieldTypeFor("memo") = dbMemo, "memo"
    TestAssert FieldTypeFor("longinteger") = dbLong, "longinteger"
    TestAssert FieldTypeFor("integer") = dbInteger, "integer"
    TestAssert FieldTypeFor("byte") = dbByte, "byte"
    TestAssert FieldTypeFor("single") = dbSingle, "single"
    TestAssert FieldTypeFor("double") = dbDouble, "double"
    TestAssert FieldTypeFor("currency") = dbCurrency, "currency"
    TestAssert FieldTypeFor("datetime") = dbDate, "datetime"
    TestAssert FieldTypeFor("yesno") = dbBoolean, "yesno"
    TestAssert FieldTypeFor("oleobject") = dbLongBinary, "oleobject"
    TestAssert FieldTypeFor("autonumber") = dbLong, "autonumber is a long"
    TestAssert FieldTypeFor("hyperlink") = dbMemo, "hyperlink is a memo"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestFieldAttributes
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Autonumber and hyperlink are ordinary types plus an attribute flag.
'---------------------------------------------------------------------------------------
'
Public Sub TestFieldAttributes()

    TestAssert HasAttribute(FieldEntry("autonumber")("Attributes"), dbAutoIncrField), _
        "autonumber sets dbAutoIncrField"
    TestAssert HasAttribute(FieldEntry("hyperlink")("Attributes"), dbHyperlinkField), _
        "hyperlink sets dbHyperlinkField"
    TestAssert FieldEntry("longinteger")("Attributes") = 0, "plain types set no attributes"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestTextFieldSizing
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Text size comes from xsd:maxLength, and falls back to the Access default
'           : when the facet is absent or nonsensical.
'---------------------------------------------------------------------------------------
'
Public Sub TestTextFieldSizing()

    Dim dField As Dictionary

    Set dField = ParseSingleField(TextFieldXml(50))
    TestAssert dField("Type") = dbText, "text type"
    TestAssert dField("Size") = 50, "size read from maxLength"

    Set dField = ParseSingleField(TextFieldXml(255))
    TestAssert dField("Size") = 255, "full-width text"

    ' No simpleType restriction at all.
    Set dField = ParseSingleField( _
        "<xsd:element name=""F"" minOccurs=""0"" od:jetType=""text""></xsd:element>")
    TestAssert dField("Size") = 255, "missing maxLength defaults to 255"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestDecimalPrecisionAndScale
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Decimal carries precision and scale from the xsd digit facets.
'---------------------------------------------------------------------------------------
'
Public Sub TestDecimalPrecisionAndScale()

    Dim dField As Dictionary

    Set dField = ParseSingleField( _
        "<xsd:element name=""F"" minOccurs=""0"" od:jetType=""decimal"">" & _
        "<xsd:simpleType><xsd:restriction base=""xsd:decimal"">" & _
        "<xsd:totalDigits value=""18""></xsd:totalDigits>" & _
        "<xsd:fractionDigits value=""4""></xsd:fractionDigits>" & _
        "</xsd:restriction></xsd:simpleType></xsd:element>")

    TestAssert dField("Type") = dbDecimal, "decimal type"
    TestAssert dField("Precision") = 18, "precision from totalDigits"
    TestAssert dField("Scale") = 4, "scale from fractionDigits"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestBigIntSentinel
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Access has no schema representation for a Large Number, and writes a bare
'           : decimal restricted to zero total digits instead. Recognizing that shape is
'           : what lets the DAO path create the field correctly rather than making a
'           : dbDecimal(38,0) and repairing it afterwards, the way the ImportXML path has
'           : to. A real decimal must not be mistaken for it.
'---------------------------------------------------------------------------------------
'
Public Sub TestBigIntSentinel()

    Dim dField As Dictionary

    Set dField = ParseSingleField( _
        "<xsd:element name=""F"" minOccurs=""0"">" & _
        "<xsd:simpleType><xsd:restriction base=""xsd:decimal"">" & _
        "<xsd:totalDigits value=""0""></xsd:totalDigits>" & _
        "<xsd:fractionDigits value=""0""></xsd:fractionDigits>" & _
        "</xsd:restriction></xsd:simpleType></xsd:element>")
    TestAssert dField("Type") = dbBigInt, "no jetType + decimal + totalDigits 0 is a Large Number"

    ' A genuine decimal declares its jetType and a non-zero precision.
    Set dField = ParseSingleField( _
        "<xsd:element name=""F"" minOccurs=""0"" od:jetType=""decimal"">" & _
        "<xsd:simpleType><xsd:restriction base=""xsd:decimal"">" & _
        "<xsd:totalDigits value=""18""></xsd:totalDigits>" & _
        "</xsd:restriction></xsd:simpleType></xsd:element>")
    TestAssert dField("Type") = dbDecimal, "a real decimal is not the sentinel"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestIndexParsing
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Index keys are a single space-separated (and space-terminated) attribute
'           : rather than nested elements, and the sort direction is a parallel list.
'---------------------------------------------------------------------------------------
'
Public Sub TestIndexParsing()

    Dim dSchema As Dictionary
    Dim dIndex As Dictionary

    Set dSchema = ParseTableXml( _
        PlainFieldXml("ID", "longinteger") & PlainFieldXml("Code", "longinteger"), _
        "<od:index index-name=""PrimaryKey"" index-key=""ID Code "" primary=""yes"" " & _
        "unique=""yes"" clustered=""no"" order=""asc asc""></od:index>")

    TestAssert dSchema("Indexes").Count = 1, "one index"
    Set dIndex = dSchema("Indexes")(1)
    TestAssert dIndex("Name") = "PrimaryKey", "index name"
    TestAssert dIndex("Primary"), "primary flag"
    TestAssert dIndex("Unique"), "unique flag"
    TestAssert dIndex("Fields").Count = 2, "trailing space does not add an empty field"
    TestAssert dIndex("Fields")(1)("Name") = "ID", "first key field"
    TestAssert dIndex("Fields")(2)("Name") = "Code", "second key field"
    TestAssert Not dIndex("Fields")(1)("Descending"), "ascending by default"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestIndexDescendingOrder
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : The order attribute is positional against index-key.
'---------------------------------------------------------------------------------------
'
Public Sub TestIndexDescendingOrder()

    Dim dIndex As Dictionary

    Set dIndex = ParseTableXml( _
        PlainFieldXml("ID", "longinteger") & PlainFieldXml("Code", "longinteger"), _
        "<od:index index-name=""Sorted"" index-key=""ID Code "" primary=""no"" " & _
        "unique=""no"" clustered=""no"" order=""asc desc""></od:index>")("Indexes")(1)

    TestAssert Not dIndex("Fields")(1)("Descending"), "first field ascending"
    TestAssert dIndex("Fields")(2)("Descending"), "second field descending"
    TestAssert Not dIndex("Primary"), "not primary"
    TestAssert Not dIndex("Unique"), "not unique"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestXmlNameDecoding
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Field and index names arrive escaped, and must come back out as the names
'           : DAO has to be given.
'---------------------------------------------------------------------------------------
'
Public Sub TestXmlNameDecoding()

    Dim dSchema As Dictionary

    Set dSchema = ParseTableXml( _
        PlainFieldXml("Index_x0026_Test", "longinteger"), _
        "<od:index index-name=""My_x0020_Index"" index-key=""Index_x0026_Test "" " & _
        "primary=""no"" unique=""no"" clustered=""no"" order=""asc""></od:index>")

    TestAssert dSchema("Fields")(1)("Name") = "Index&Test", "field name decoded"
    TestAssert dSchema("Indexes")(1)("Name") = "My Index", "index name decoded"
    TestAssert dSchema("Indexes")(1)("Fields")(1)("Name") = "Index&Test", "index key decoded"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestUnescapeXmlNameRoundTrip
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : UnescapeXmlName must invert EscapeXmlName, including the awkward cases
'           : that motivated the "escape an underscore only before a lowercase x" rule.
'---------------------------------------------------------------------------------------
'
Public Sub TestUnescapeXmlNameRoundTrip()

    TestAssert UnescapeXmlName(EscapeXmlName("Plain")) = "Plain", "plain name"
    TestAssert UnescapeXmlName(EscapeXmlName("Has Space")) = "Has Space", "space"
    TestAssert UnescapeXmlName(EscapeXmlName("Index&Test")) = "Index&Test", "ampersand"
    TestAssert UnescapeXmlName(EscapeXmlName("NotReq'd")) = "NotReq'd", "apostrophe"
    TestAssert UnescapeXmlName(EscapeXmlName("a_x1")) = "a_x1", "underscore before lowercase x"
    TestAssert UnescapeXmlName(EscapeXmlName("a_x005F_b")) = "a_x005F_b", "literal escape sequence"
    TestAssert UnescapeXmlName(EscapeXmlName("a_Xb")) = "a_Xb", "uppercase X is not escaped"

    ' Text that merely resembles an escape is left alone.
    TestAssert UnescapeXmlName("a_xZZZZ_b") = "a_xZZZZ_b", "invalid hex is literal"
    TestAssert UnescapeXmlName("a_x12_b") = "a_x12_b", "too few hex digits is literal"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestPropertyParsing
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Field and table properties keep their declared DAO type, and the value is
'           : converted to match it.
'---------------------------------------------------------------------------------------
'
Public Sub TestPropertyParsing()

    Dim dSchema As Dictionary
    Dim dProps As Dictionary

    Set dSchema = ParseTableXml( _
        "<xsd:element name=""F"" minOccurs=""0"" od:jetType=""text"">" & _
        "<xsd:annotation><xsd:appinfo>" & _
        "<od:fieldProperty name=""Required"" type=""1"" value=""1""></od:fieldProperty>" & _
        "<od:fieldProperty name=""ColumnWidth"" type=""3"" value=""-1""></od:fieldProperty>" & _
        "<od:fieldProperty name=""Caption"" type=""10"" value=""My Field""></od:fieldProperty>" & _
        "<od:fieldProperty name=""BackTint"" type=""6"" value=""100""></od:fieldProperty>" & _
        "</xsd:appinfo></xsd:annotation></xsd:element>", _
        "<od:tableProperty name=""Description"" type=""10"" value=""A table""></od:tableProperty>")

    Set dProps = dSchema("Fields")(1)("Properties")
    TestAssert dProps("Required")("Value") = True, "boolean 1 is True"
    TestAssert dProps("ColumnWidth")("Value") = -1, "signed integer"
    TestAssert dProps("ColumnWidth")("Type") = 3, "declared type preserved"
    TestAssert dProps("Caption")("Value") = "My Field", "string value"
    TestAssert dProps("BackTint")("Value") = 100, "single value"

    TestAssert dSchema("Properties")("Description")("Value") = "A table", "table property"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestRequiredIsNotInferredFromNonNullable
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : An autonumber primary key is written with od:nonNullable="yes" and no
'           : Required property. Inferring Required from it would add a property the
'           : original table never had, and the verification re-export would stop
'           : matching. This pins that decision.
'---------------------------------------------------------------------------------------
'
Public Sub TestRequiredIsNotInferredFromNonNullable()

    Dim dProps As Dictionary

    Set dProps = ParseSingleField( _
        "<xsd:element name=""ID"" minOccurs=""1"" od:jetType=""autonumber"" " & _
        "od:sqlSType=""int"" od:autoUnique=""yes"" od:nonNullable=""yes"" " & _
        "type=""xsd:int""></xsd:element>")("Properties")

    TestAssert Not dProps.Exists("Required"), "nonNullable does not create a Required property"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestUnsupportedConstructsDecline
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Every construct we refuse must produce Nothing and a reason, not a partial
'           : model. This is the guard that keeps a silently incomplete table from ever
'           : reaching the database.
'---------------------------------------------------------------------------------------
'
Public Sub TestUnsupportedConstructsDecline()

    ' Attachment / multi-value field.
    TestAssert Not Parses( _
        "<xsd:element name=""F"" minOccurs=""0"" od:jetType=""complex"" " & _
        "od:jetComplexType=""MSysComplexType_Attachment"" maxOccurs=""unbounded""></xsd:element>"), _
        "complex field declines"

    ' Calculated column.
    TestAssert Not Parses( _
        "<xsd:element name=""F"" minOccurs=""0"" od:jetType=""text"" " & _
        "od:expression=""[A]+[B]""></xsd:element>"), _
        "calculated field declines"

    ' A jetType from some future version of Access.
    TestAssert Not Parses( _
        "<xsd:element name=""F"" minOccurs=""0"" od:jetType=""quantum""></xsd:element>"), _
        "unknown jetType declines"

    ' No jetType and nothing that looks like the Large Number sentinel.
    TestAssert Not Parses( _
        "<xsd:element name=""F"" minOccurs=""0"" type=""xsd:string""></xsd:element>"), _
        "untyped field declines"

    ' An od: attribute we have never seen.
    TestAssert Not Parses( _
        "<xsd:element name=""F"" minOccurs=""0"" od:jetType=""text"" " & _
        "od:someNewThing=""1""></xsd:element>"), _
        "unknown od: attribute declines"

    TestAssert Len(GetLastDeclineReason()) > 0, "a decline records a reason"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestUnknownElementDeclines
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Unknown vocabulary anywhere in the document, not just on a field.
'---------------------------------------------------------------------------------------
'
Public Sub TestUnknownElementDeclines()

    TestAssert Not Parses(PlainFieldXml("ID", "longinteger"), _
        "<od:futureThing name=""x""></od:futureThing>"), _
        "unknown od: element declines"

    TestAssert GetTableDefSchemaFromXml("this is not xml") Is Nothing, _
        "malformed XML declines"

    TestAssert GetTableDefSchemaFromXml(vbNullString) Is Nothing, _
        "empty content declines"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestKnownGoodTableParses
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : A whole ordinary table, to confirm the pieces compose.
'---------------------------------------------------------------------------------------
'
Public Sub TestKnownGoodTableParses()

    Dim dSchema As Dictionary

    Set dSchema = ParseTableXml( _
        "<xsd:element name=""ID"" minOccurs=""1"" od:jetType=""autonumber"" " & _
            "od:sqlSType=""int"" od:autoUnique=""yes"" od:nonNullable=""yes"" " & _
            "type=""xsd:int""></xsd:element>" & _
        TextFieldXml(255) & _
        PlainFieldXml("Created", "datetime"), _
        "<od:index index-name=""PrimaryKey"" index-key=""ID "" primary=""yes"" " & _
        "unique=""yes"" clustered=""no"" order=""asc""></od:index>")

    TestAssert Not dSchema Is Nothing, "parses"
    TestAssert dSchema("Name") = "TestTable", "table name"
    TestAssert dSchema("Fields").Count = 3, "three fields"
    TestAssert dSchema("Indexes").Count = 1, "one index"
    TestAssert Len(GetLastDeclineReason()) = 0, "no decline reason on success"

End Sub


' --- Helpers (parameterized, so they are not discovered as tests) -----------------------


'---------------------------------------------------------------------------------------
' Procedure : WrapTableXml
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Wrap field and appinfo fragments in the envelope Access exports, so each
'           : test only has to state the part it cares about.
'---------------------------------------------------------------------------------------
'
Private Function WrapTableXml(strFields As String, strTableAppInfo As String) As String

    Dim strAnnotation As String

    If Len(strTableAppInfo) > 0 Then
        strAnnotation = "<xsd:annotation><xsd:appinfo>" & strTableAppInfo & _
            "</xsd:appinfo></xsd:annotation>"
    End If

    WrapTableXml = "<?xml version=""1.0""?>" & _
        "<xsd:schema xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" " & _
        "xmlns:od=""urn:schemas-microsoft-com:officedata"">" & _
        "<xsd:element name=""dataroot""><xsd:complexType><xsd:sequence>" & _
        "<xsd:element ref=""TestTable"" minOccurs=""0"" maxOccurs=""unbounded""></xsd:element>" & _
        "</xsd:sequence>" & _
        "<xsd:attribute name=""generated"" type=""xsd:dateTime""></xsd:attribute>" & _
        "</xsd:complexType></xsd:element>" & _
        "<xsd:element name=""TestTable"">" & strAnnotation & _
        "<xsd:complexType><xsd:sequence>" & strFields & _
        "</xsd:sequence></xsd:complexType></xsd:element></xsd:schema>"

End Function


Private Function ParseTableXml(strFields As String, _
    Optional strTableAppInfo As String = vbNullString) As Dictionary
    Set ParseTableXml = GetTableDefSchemaFromXml(WrapTableXml(strFields, strTableAppInfo))
End Function


Private Function Parses(strFields As String, _
    Optional strTableAppInfo As String = vbNullString) As Boolean
    Parses = Not (ParseTableXml(strFields, strTableAppInfo) Is Nothing)
End Function


Private Function ParseSingleField(strFieldXml As String) As Dictionary
    Set ParseSingleField = ParseTableXml(strFieldXml)("Fields")(1)
End Function


Private Function FieldEntry(strJetType As String) As Dictionary
    Set FieldEntry = ParseSingleField(PlainFieldXml("F", strJetType))
End Function


Private Function FieldTypeFor(strJetType As String) As Long
    FieldTypeFor = FieldEntry(strJetType)("Type")
End Function


Private Function PlainFieldXml(strName As String, strJetType As String) As String
    PlainFieldXml = "<xsd:element name=""" & strName & """ minOccurs=""0"" od:jetType=""" & _
        strJetType & """></xsd:element>"
End Function


'---------------------------------------------------------------------------------------
' Procedure : HasAttribute
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Bit test for a field's Attributes flags.
'           :
'           : This is a helper rather than an inline expression because a Sub call's
'           : argument list cannot begin with "(" -- VBA reads the parenthesized group as
'           : the whole argument and then fails on what follows. ByVal because the
'           : attributes arrive from a Dictionary as a Variant.
'---------------------------------------------------------------------------------------
'
Private Function HasAttribute(ByVal lngAttributes As Long, ByVal lngFlag As Long) As Boolean
    HasAttribute = ((lngAttributes And lngFlag) = lngFlag)
End Function


Private Function TextFieldXml(ByVal lngSize As Long) As String
    TextFieldXml = "<xsd:element name=""Name"" minOccurs=""0"" od:jetType=""text"" " & _
        "od:sqlSType=""nvarchar"">" & _
        "<xsd:simpleType><xsd:restriction base=""xsd:string"">" & _
        "<xsd:maxLength value=""" & lngSize & """></xsd:maxLength>" & _
        "</xsd:restriction></xsd:simpleType></xsd:element>"
End Function
