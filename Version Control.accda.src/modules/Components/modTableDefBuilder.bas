Attribute VB_Name = "modTableDefBuilder"
'---------------------------------------------------------------------------------------
' Module    : modTableDefBuilder
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Build a local table definition directly through DAO from an exported
'           : tbldefs XML file, as an alternative to Application.ImportXML.
'           :
'           : Why this exists: ImportXML with acStructureOnly costs 276-281 seconds in a
'           : database holding ~3,700 saved queries, against 0.02 seconds for the same
'           : file in a small one. Isolation testing showed the cost tracks the number of
'           : table references across the saved-query catalog, not the table being
'           : imported, so it cannot be tuned from our side -- the only remedy is not to
'           : make the call. See DECISIONS.md 2026-07-30.
'           :
'           : Safety model. This module is deliberately unwilling to guess. Parsing walks
'           : the whole document and rejects any element or attribute it does not
'           : explicitly recognize, so a construct we have never seen causes a clean
'           : fallback to ImportXML rather than a table that is silently missing
'           : something. The caller then re-exports the finished table and compares it
'           : against the source file, so even a construct we recognize but reproduce
'           : imperfectly is caught before it can be committed.
'           :
'           : Property names are the one thing NOT whitelisted, because the mechanism for
'           : them is generic: everything except a handful of native DAO members goes
'           : through SetDAOProperty, which is the same path Access itself uses and works
'           : for names we have never seen. A native member we failed to route would
'           : either raise (caught here) or fail verification (caught by the caller).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Components")

Private Const ModuleName As String = "modTableDefBuilder"

' XML namespaces used by Application.ExportXML
Private Const NS_XSD As String = "http://www.w3.org/2001/XMLSchema"
Private Const NS_OD As String = "urn:schemas-microsoft-com:officedata"
Private Const NS_XMLNS As String = "http://www.w3.org/2000/xmlns/"

' Access encodes a memo field as a string restricted to this length. Used to tell a
' genuine text field from a memo when the schema carries no other signal.
Private Const MAX_TEXT_SIZE As Long = 255

' Reason the last parse or build declined, for debug logging by the caller.
Private m_strDeclineReason As String

' True once a call got as far as touching the database. Lets the caller charge its
' circuit breaker for work that actually cost something, rather than for a parse that
' declined an unsupported construct in a millisecond.
Private m_blnAttemptedBuild As Boolean

' Test hook. The caller only opens the eligibility gate on databases large enough for
' ImportXML to hurt, which no test database is. Setting this bypasses that gate so the
' round-trip fixtures exercise this path. Not used in production code.
Public FastPathTestOverride As Boolean


'---------------------------------------------------------------------------------------
' Procedure : TryBuildTableFromDefXml
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Create the table described by an exported tbldefs XML file using DAO.
'           : Returns true only when the table was fully created. On any failure the
'           : table is removed again, so the caller can fall back to Application.ImportXML
'           : without colliding with a partial object. (A collision would be silent --
'           : ImportXML would import as "MyTable1" rather than raise.)
'---------------------------------------------------------------------------------------
'
Public Function TryBuildTableFromDefXml(strXmlFile As String, strTableName As String) As Boolean

    Dim dSchema As Dictionary
    Dim blnCreated As Boolean

    m_strDeclineReason = vbNullString
    m_blnAttemptedBuild = False

    If Not FSO.FileExists(strXmlFile) Then
        m_strDeclineReason = "source file not found"
        Exit Function
    End If

    ' Never build over an existing object. The merge path deletes first; anything else
    ' reaching here means the caller's assumptions do not hold.
    If TableExists(strTableName) Then
        m_strDeclineReason = "table already exists"
        Exit Function
    End If

    Set dSchema = GetTableDefSchemaFromXml(ReadFile(strXmlFile))
    If dSchema Is Nothing Then Exit Function

    TryBuildTableFromDefXml = CreateTableFromSchema(dSchema, strTableName, blnCreated)

    ' Remove any partial object left by a failure part way through.
    If Not TryBuildTableFromDefXml And blnCreated Then RemovePartialTable strTableName

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetLastDeclineReason
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Why the last call declined to build. Empty when it succeeded.
'           :
'           : "Succeeded" means the caller kept the table, not merely that this module
'           : produced one -- see RecordVerificationFailure.
'---------------------------------------------------------------------------------------
'
Public Function GetLastDeclineReason() As String
    GetLastDeclineReason = m_strDeclineReason
End Function


'---------------------------------------------------------------------------------------
' Procedure : RecordVerificationFailure
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Let the caller report that it rejected and discarded a table this module
'           : built, so GetLastDeclineReason describes the outcome of the import rather
'           : than only the outcome of the parse.
'           :
'           : Without this, a build that parses cleanly but fails the caller's re-export
'           : comparison leaves no decline reason, and anything reading that reason as a
'           : proxy for "the DAO path was used" is misled. The round-trip harness read it
'           : exactly that way and passed a fixture that had silently fallen back to
'           : Application.ImportXML.
'---------------------------------------------------------------------------------------
'
Public Sub RecordVerificationFailure(strReason As String)
    m_strDeclineReason = strReason
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetLastAttemptReachedBuild
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : True when the last call got past parsing and started creating the table.
'           : A failure here cost real time; a parse decline did not.
'---------------------------------------------------------------------------------------
'
Public Function GetLastAttemptReachedBuild() As Boolean
    GetLastAttemptReachedBuild = m_blnAttemptedBuild
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTableDefSchemaFromXml
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Parse exported table definition XML into a schema model, or return Nothing
'           : when the document contains anything this module does not fully understand.
'           : Exposed (in-project) so tests can assert the model without a database.
'           :
'           : Model shape:
'           :   Name       -> table name
'           :   Fields     -> Collection of field dictionaries
'           :                   Name, Type, Size, Attributes, Precision, Scale,
'           :                   Properties -> name -> {Type, Value}
'           :   Indexes    -> Collection of index dictionaries
'           :                   Name, Primary, Unique,
'           :                   Fields -> Collection of {Name, Descending}
'           :   Properties -> name -> {Type, Value}
'---------------------------------------------------------------------------------------
'
Public Function GetTableDefSchemaFromXml(strXml As String) As Dictionary

    Dim objXml As MSXML2.DOMDocument60
    Dim objTable As MSXML2.IXMLDOMNode
    Dim dSchema As Dictionary

    m_strDeclineReason = vbNullString
    If Len(strXml) = 0 Then
        m_strDeclineReason = "empty source file"
        Exit Function
    End If

    ' Own the error state rather than inheriting the caller's On Error Resume Next,
    ' which would let a parse failure run on into the builder with a partial model.
    LogUnhandledErrors
    On Error GoTo ErrHandler

    Perf.OperationStart "Parse Table Def XML"

    Set objXml = New MSXML2.DOMDocument60
    objXml.async = False
    If objXml.LoadXML(strXml) Then
        ' Reject unknown vocabulary before reading anything, so a document we only
        ' partly understand never reaches the builder.
        If ValidateVocabulary(objXml.documentElement) Then
            Set objTable = GetTableElement(objXml)
            If objTable Is Nothing Then
                m_strDeclineReason = "no table element found"
            Else
                Set dSchema = ParseTableElement(objTable)
            End If
        End If
    Else
        m_strDeclineReason = "malformed XML"
    End If

    Perf.OperationEnd

    Set GetTableDefSchemaFromXml = dSchema
    Exit Function

ErrHandler:
    m_strDeclineReason = "parse failed: " & Err.Number & " " & Err.Description
    Err.Clear
    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : ValidateVocabulary
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Walk the document and confirm every element and attribute is one we know
'           : how to reproduce. This is the guard that turns "a construct we have never
'           : seen" into a fallback rather than a quietly incomplete table.
'---------------------------------------------------------------------------------------
'
Private Function ValidateVocabulary(ByVal objNode As MSXML2.IXMLDOMNode) As Boolean

    Dim objChild As MSXML2.IXMLDOMNode
    Dim objAttr As MSXML2.IXMLDOMNode
    Dim strLocal As String
    Dim strNs As String

    If objNode Is Nothing Then
        m_strDeclineReason = "empty document"
        Exit Function
    End If

    strNs = objNode.NamespaceURI
    strLocal = objNode.baseName

    Select Case strNs
        Case NS_XSD
            Select Case strLocal
                Case "schema", "element", "complexType", "sequence", "attribute", _
                     "annotation", "appinfo", "simpleType", "restriction", _
                     "maxLength", "totalDigits", "fractionDigits"
                    ' Recognized
                Case Else
                    m_strDeclineReason = "unsupported schema element 'xsd:" & strLocal & "'"
                    Exit Function
            End Select
        Case NS_OD
            Select Case strLocal
                Case "index", "tableProperty", "fieldProperty"
                    ' Recognized
                Case Else
                    m_strDeclineReason = "unsupported element 'od:" & strLocal & "'"
                    Exit Function
            End Select
        Case Else
            m_strDeclineReason = "unsupported element namespace for '" & strLocal & "'"
            Exit Function
    End Select

    ' Check attributes on this element
    For Each objAttr In objNode.Attributes
        If Not IsKnownAttribute(strNs, strLocal, objAttr) Then Exit Function
    Next objAttr

    ' Recurse into child elements
    For Each objChild In objNode.ChildNodes
        If objChild.NodeType = NODE_ELEMENT Then
            If Not ValidateVocabulary(objChild) Then Exit Function
        End If
    Next objChild

    ValidateVocabulary = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsKnownAttribute
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Returns true if this attribute is one we read (or can safely ignore) on the
'           : given element. Also rejects the od: values that mark a construct DAO cannot
'           : reproduce here -- complex/multi-value fields and calculated columns.
'---------------------------------------------------------------------------------------
'
Private Function IsKnownAttribute(strElementNs As String, strElement As String, _
    ByVal objAttr As MSXML2.IXMLDOMNode) As Boolean

    Dim strLocal As String
    Dim strNs As String

    strNs = objAttr.NamespaceURI
    strLocal = objAttr.baseName

    ' Namespace declarations carry no data.
    If strNs = NS_XMLNS Or StrComp(strLocal, "xmlns", vbTextCompare) = 0 Then
        IsKnownAttribute = True
        Exit Function
    End If

    If strNs = NS_OD Then
        Select Case strLocal
            Case "jetType"
                ' Attachments and multi-value fields are not reproducible through
                ' CreateField, so hand the whole table back to ImportXML.
                If StrComp(objAttr.Text, "complex", vbTextCompare) = 0 Then
                    m_strDeclineReason = "complex (attachment or multi-value) field"
                    Exit Function
                End If
            Case "jetComplexType"
                m_strDeclineReason = "complex field type '" & objAttr.Text & "'"
                Exit Function
            Case "expression"
                m_strDeclineReason = "calculated field"
                Exit Function
            Case "sqlSType", "autoUnique", "nonNullable", "hyperlink"
                ' Recognized
            Case Else
                m_strDeclineReason = "unsupported attribute 'od:" & strLocal & "'"
                Exit Function
        End Select
        IsKnownAttribute = True
        Exit Function
    End If

    ' Unprefixed attributes, by owning element
    If strElementNs = NS_OD Then
        Select Case strElement
            Case "index"
                Select Case strLocal
                    Case "index-name", "index-key", "primary", "unique", "clustered", "order"
                        IsKnownAttribute = True
                End Select
            Case "tableProperty", "fieldProperty"
                Select Case strLocal
                    Case "name", "type", "value"
                        IsKnownAttribute = True
                End Select
        End Select
    Else
        Select Case strElement
            Case "element"
                Select Case strLocal
                    Case "name", "ref", "minOccurs", "maxOccurs", "type"
                        IsKnownAttribute = True
                End Select
            Case "attribute"
                Select Case strLocal
                    Case "name", "type"
                        IsKnownAttribute = True
                End Select
            Case "restriction"
                If strLocal = "base" Then IsKnownAttribute = True
            Case "maxLength", "totalDigits", "fractionDigits"
                If strLocal = "value" Then IsKnownAttribute = True
            Case "schema", "complexType", "sequence", "annotation", "appinfo"
                ' No attributes expected, but none of these carry field data.
        End Select
    End If

    If Not IsKnownAttribute Then
        m_strDeclineReason = "unsupported attribute '" & strLocal & "' on '" & strElement & "'"
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTableElement
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Return the xsd:element that describes the table. An export contains two
'           : top-level elements: the "dataroot" envelope and the table itself.
'---------------------------------------------------------------------------------------
'
Private Function GetTableElement(ByVal objXml As MSXML2.DOMDocument60) As MSXML2.IXMLDOMNode

    Dim objChild As MSXML2.IXMLDOMNode
    Dim objName As MSXML2.IXMLDOMNode

    For Each objChild In objXml.documentElement.ChildNodes
        If objChild.NodeType = NODE_ELEMENT Then
            If objChild.NamespaceURI = NS_XSD And objChild.baseName = "element" Then
                Set objName = objChild.Attributes.getNamedItem("name")
                If Not objName Is Nothing Then
                    If StrComp(objName.Text, "dataroot", vbTextCompare) <> 0 Then
                        Set GetTableElement = objChild
                        Exit Function
                    End If
                End If
            End If
        End If
    Next objChild

End Function


'---------------------------------------------------------------------------------------
' Procedure : ParseTableElement
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Build the schema model from the table's xsd:element node.
'---------------------------------------------------------------------------------------
'
Private Function ParseTableElement(ByVal objTable As MSXML2.IXMLDOMNode) As Dictionary

    Dim dSchema As Dictionary
    Dim colFields As Collection
    Dim objNode As MSXML2.IXMLDOMNode
    Dim objSequence As MSXML2.IXMLDOMNode
    Dim dField As Dictionary

    Set dSchema = New Dictionary
    dSchema.Add "Name", UnescapeXmlName(GetAttr(objTable, "name"))
    dSchema.Add "Indexes", ParseIndexes(objTable)
    dSchema.Add "Properties", ParseProperties(objTable, "tableProperty")

    Set colFields = New Collection
    Set objSequence = FindDescendant(objTable, NS_XSD, "sequence")
    If objSequence Is Nothing Then
        m_strDeclineReason = "no field sequence found"
        Exit Function
    End If

    For Each objNode In objSequence.ChildNodes
        If objNode.NodeType = NODE_ELEMENT Then
            If objNode.NamespaceURI = NS_XSD And objNode.baseName = "element" Then
                Set dField = ParseFieldElement(objNode)
                If dField Is Nothing Then Exit Function
                colFields.Add dField
            End If
        End If
    Next objNode

    If colFields.Count = 0 Then
        m_strDeclineReason = "no fields found"
        Exit Function
    End If

    dSchema.Add "Fields", colFields
    Set ParseTableElement = dSchema

End Function


'---------------------------------------------------------------------------------------
' Procedure : ParseFieldElement
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Translate one field's xsd:element into name, DAO type, size and properties.
'           :
'           : The type map is the piece with no existing counterpart in this codebase.
'           : Note the Big Integer case: ExportXML has no schema representation for
'           : dbBigInt and writes a bare xsd:decimal restricted to totalDigits="0",
'           : which is why ImportXML mis-creates those fields as dbDecimal(38,0) and
'           : clsDbTableDef has to repair them afterwards. Recognizing the sentinel here
'           : creates the field correctly the first time.
'---------------------------------------------------------------------------------------
'
Private Function ParseFieldElement(ByVal objField As MSXML2.IXMLDOMNode) As Dictionary

    Dim dField As Dictionary
    Dim strJetType As String
    Dim lngType As Long
    Dim lngSize As Long
    Dim lngAttributes As Long
    Dim lngPrecision As Long
    Dim lngScale As Long
    Dim strName As String

    strName = UnescapeXmlName(GetAttr(objField, "name"))
    strJetType = GetAttrNs(objField, NS_OD, "jetType")

    Select Case LCase$(strJetType)

        Case "text"
            lngType = dbText
            lngSize = GetFacetValue(objField, "maxLength")
            If lngSize <= 0 Or lngSize > MAX_TEXT_SIZE Then lngSize = MAX_TEXT_SIZE

        Case "memo"
            lngType = dbMemo

        Case "hyperlink"
            ' Stored as a memo carrying the hyperlink attribute.
            lngType = dbMemo
            lngAttributes = dbHyperlinkField

        Case "autonumber"
            lngType = dbLong
            lngAttributes = dbAutoIncrField

        Case "longinteger":     lngType = dbLong
        Case "integer":         lngType = dbInteger
        Case "byte":            lngType = dbByte
        Case "single":          lngType = dbSingle
        Case "double":          lngType = dbDouble
        Case "currency":        lngType = dbCurrency
        Case "datetime":        lngType = dbDate
        Case "yesno":           lngType = dbBoolean
        Case "oleobject":       lngType = dbLongBinary

        Case "decimal"
            lngType = dbDecimal
            lngPrecision = GetFacetValue(objField, "totalDigits")
            lngScale = GetFacetValue(objField, "fractionDigits")

        Case vbNullString
            ' No jetType at all. The only shape we accept is the Big Integer sentinel.
            If IsBigIntSentinel(objField) Then
                lngType = dbBigInt
            Else
                m_strDeclineReason = "field '" & strName & "' has no recognizable type"
                Exit Function
            End If

        Case Else
            m_strDeclineReason = "unsupported field type '" & strJetType & _
                "' on field '" & strName & "'"
            Exit Function

    End Select

    Set dField = New Dictionary
    dField.Add "Name", strName
    dField.Add "Type", lngType
    dField.Add "Size", lngSize
    dField.Add "Attributes", lngAttributes
    dField.Add "Precision", lngPrecision
    dField.Add "Scale", lngScale
    dField.Add "Properties", ParseProperties(objField, "fieldProperty")

    ' Note what is deliberately NOT read here: od:nonNullable. On an autonumber primary
    ' key Access writes nonNullable="yes" with no Required field property at all, because
    ' the non-nullability comes from the primary index rather than from Required. Setting
    ' Required from it would add a property the original table does not have, and the
    ' re-export would no longer match the source. Required comes only from its own
    ' fieldProperty entry.

    Set ParseFieldElement = dField

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsBigIntSentinel
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Returns true for the shape ExportXML writes in place of a Big Integer:
'           : a decimal restriction with totalDigits="0". A real decimal always carries
'           : od:jetType="decimal" and a non-zero precision.
'---------------------------------------------------------------------------------------
'
Private Function IsBigIntSentinel(ByVal objField As MSXML2.IXMLDOMNode) As Boolean

    Dim objRestriction As MSXML2.IXMLDOMNode

    Set objRestriction = FindDescendant(objField, NS_XSD, "restriction")
    If objRestriction Is Nothing Then Exit Function
    If StrComp(GetAttr(objRestriction, "base"), "xsd:decimal", vbTextCompare) <> 0 Then Exit Function

    IsBigIntSentinel = (GetFacetValue(objField, "totalDigits") = 0)

End Function


'---------------------------------------------------------------------------------------
' Procedure : ParseIndexes
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Read the od:index elements. Indexes are flat attributes rather than nested
'           : field elements: index-key holds the field names separated (and terminated)
'           : by spaces, and order holds one direction per field.
'---------------------------------------------------------------------------------------
'
Private Function ParseIndexes(ByVal objTable As MSXML2.IXMLDOMNode) As Collection

    Dim colIndexes As Collection
    Dim objAppInfo As MSXML2.IXMLDOMNode
    Dim objNode As MSXML2.IXMLDOMNode
    Dim dIndex As Dictionary
    Dim colIdxFields As Collection
    Dim dIdxField As Dictionary
    Dim varKeys As Variant
    Dim varOrders As Variant
    Dim lngPos As Long
    Dim lngField As Long

    Set colIndexes = New Collection
    Set ParseIndexes = colIndexes

    Set objAppInfo = FindDescendant(objTable, NS_XSD, "appinfo")
    If objAppInfo Is Nothing Then Exit Function

    For Each objNode In objAppInfo.ChildNodes
        If objNode.NodeType = NODE_ELEMENT Then
            If objNode.NamespaceURI = NS_OD And objNode.baseName = "index" Then

                varKeys = Split(GetAttr(objNode, "index-key"), " ")
                varOrders = Split(GetAttr(objNode, "order"), " ")

                Set colIdxFields = New Collection
                lngField = 0
                For lngPos = LBound(varKeys) To UBound(varKeys)
                    If Len(Trim$(varKeys(lngPos))) > 0 Then
                        Set dIdxField = New Dictionary
                        dIdxField.Add "Name", UnescapeXmlName(CStr(varKeys(lngPos)))
                        dIdxField.Add "Descending", IsDescending(varOrders, lngField)
                        colIdxFields.Add dIdxField
                        lngField = lngField + 1
                    End If
                Next lngPos

                If colIdxFields.Count > 0 Then
                    Set dIndex = New Dictionary
                    dIndex.Add "Name", UnescapeXmlName(GetAttr(objNode, "index-name"))
                    dIndex.Add "Primary", IsYes(GetAttr(objNode, "primary"))
                    dIndex.Add "Unique", IsYes(GetAttr(objNode, "unique"))
                    dIndex.Add "Fields", colIdxFields
                    colIndexes.Add dIndex
                End If

            End If
        End If
    Next objNode

End Function


'---------------------------------------------------------------------------------------
' Procedure : ParseProperties
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Read od:tableProperty or od:fieldProperty children into a dictionary of
'           : name -> {Type, Value}, matching the shape SetDAOProperty already consumes
'           : for linked tables.
'---------------------------------------------------------------------------------------
'
Private Function ParseProperties(ByVal objParent As MSXML2.IXMLDOMNode, ByVal strPropElement As String) As Dictionary

    Dim dProps As Dictionary
    Dim objAppInfo As MSXML2.IXMLDOMNode
    Dim objNode As MSXML2.IXMLDOMNode
    Dim strName As String
    Dim intType As Integer

    Set dProps = New Dictionary
    Set ParseProperties = dProps

    Set objAppInfo = FindDescendant(objParent, NS_XSD, "appinfo")
    If objAppInfo Is Nothing Then Exit Function

    For Each objNode In objAppInfo.ChildNodes
        If objNode.NodeType = NODE_ELEMENT Then
            If objNode.NamespaceURI = NS_OD And objNode.baseName = strPropElement Then
                strName = GetAttr(objNode, "name")
                If Len(strName) > 0 And Not dProps.Exists(strName) Then
                    intType = CInt(Val(GetAttr(objNode, "type")))
                    dProps.Add strName, BuildPropertyEntry(intType, _
                        CoercePropertyValue(intType, GetAttr(objNode, "value")))
                End If
            End If
        End If
    Next objNode

End Function


'---------------------------------------------------------------------------------------
' Procedure : CoercePropertyValue
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Convert an XML property value string to the DAO type it declares. Val() is
'           : used rather than CLng/CDbl on the raw string because XML always writes "."
'           : as the decimal separator regardless of the user's locale.
'---------------------------------------------------------------------------------------
'
Private Function CoercePropertyValue(intType As Integer, strValue As String) As Variant

    Select Case intType
        Case dbBoolean
            CoercePropertyValue = (Val(strValue) <> 0) _
                Or (StrComp(strValue, "true", vbTextCompare) = 0) _
                Or (StrComp(strValue, "yes", vbTextCompare) = 0)
        Case dbByte:                        CoercePropertyValue = CByte(Val(strValue))
        Case dbInteger:                     CoercePropertyValue = CInt(Val(strValue))
        Case dbLong:                        CoercePropertyValue = CLng(Val(strValue))
        Case dbSingle:                      CoercePropertyValue = CSng(Val(strValue))
        Case dbDouble:                      CoercePropertyValue = CDbl(Val(strValue))
        Case dbCurrency:                    CoercePropertyValue = CCur(Val(strValue))
        Case Else:                          CoercePropertyValue = strValue
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : CreateTableFromSchema
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Materialize the parsed schema as a real table.
'           :
'           : Order matters. Fields and indexes are built on the unsaved TableDef, then
'           : the whole thing is appended in one go; custom properties can only be
'           : created once the object exists, so they follow. blnCreated tells the caller
'           : whether a cleanup delete is needed if we fail after the append.
'---------------------------------------------------------------------------------------
'
Private Function CreateTableFromSchema(ByVal dSchema As Dictionary, strTableName As String, _
    ByRef blnCreated As Boolean) As Boolean

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim varField As Variant
    Dim varIndex As Variant

    LogUnhandledErrors
    On Error GoTo ErrHandler

    Perf.OperationStart "Create Table (DAO)"
    m_blnAttemptedBuild = True

    Set dbs = CurrentDb
    Set tdf = dbs.CreateTableDef(strTableName)

    For Each varField In dSchema("Fields")
        AppendField tdf, varField
    Next varField

    For Each varIndex In dSchema("Indexes")
        AppendIndex tdf, varIndex
    Next varIndex

    dbs.TableDefs.Append tdf
    blnCreated = True
    dbs.TableDefs.Refresh
    ' Drop the SharedDb cache so the verification export (and any later SharedDb
    ' caller) sees the table just appended through this local CurrentDb handle.
    ReleaseDbReferences

    ' Re-fetch through a fresh handle. The appended TableDef reference does not reliably
    ' expose the saved object's property collections.
    Set dbs = CurrentDb
    Set tdf = dbs.TableDefs(strTableName)

    For Each varField In dSchema("Fields")
        ApplyFieldProperties tdf.Fields(CStr(varField("Name"))), varField("Properties")
    Next varField

    ApplyTableProperties tdf, dSchema("Properties")

    Perf.OperationEnd
    CreateTableFromSchema = True
    Exit Function

ErrHandler:
    m_strDeclineReason = "DAO build failed: " & Err.Number & " " & Err.Description
    Err.Clear
    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : AppendField
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Create one field on the unsaved TableDef. Size, precision, attributes and
'           : the native DAO members must all be set before the field is appended.
'---------------------------------------------------------------------------------------
'
' Note on ByVal: the callers walk Collections of Dictionary, so every argument arrives as
' a Variant. A ByRef Dictionary parameter refuses that outright ("ByRef argument type
' mismatch"), because VBA will not widen a Variant into a typed reference in place. ByVal
' lets it coerce, and costs nothing -- the object reference is still shared.
Private Sub AppendField(tdf As DAO.TableDef, ByVal dField As Dictionary)

    Dim fld As DAO.Field
    Dim lngType As Long

    lngType = dField("Type")
    Set fld = tdf.CreateField(CStr(dField("Name")), lngType)

    If lngType = dbText Then fld.Size = dField("Size")

    If lngType = dbDecimal Then
        If dField("Precision") > 0 Then fld.Precision = dField("Precision")
        fld.NumericScale = dField("Scale")
    End If

    ' Combine rather than assign, so the engine's own storage flags survive.
    If dField("Attributes") <> 0 Then
        fld.Attributes = fld.Attributes Or dField("Attributes")
    End If

    ApplyNativeFieldMembers fld, dField("Properties")

    tdf.Fields.Append fld

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ApplyNativeFieldMembers
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Apply the field properties that are built-in DAO Field members rather than
'           : entries in the Properties collection. Routing these through CreateProperty
'           : instead would either raise or create a shadow property that Access ignores,
'           : so they are handled here and skipped by ApplyFieldProperties.
'---------------------------------------------------------------------------------------
'
Private Sub ApplyNativeFieldMembers(fld As DAO.Field, ByVal dProps As Dictionary)

    ' AllowZeroLength must precede Required: setting Required on a text field that does
    ' not yet allow zero-length strings is what Access itself does in that order.
    If dProps.Exists("AllowZeroLength") Then fld.AllowZeroLength = dProps("AllowZeroLength")("Value")
    If dProps.Exists("Required") Then fld.Required = dProps("Required")("Value")
    If dProps.Exists("DefaultValue") Then fld.DefaultValue = dProps("DefaultValue")("Value")
    If dProps.Exists("ValidationRule") Then fld.ValidationRule = dProps("ValidationRule")("Value")
    If dProps.Exists("ValidationText") Then fld.ValidationText = dProps("ValidationText")("Value")

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ApplyFieldProperties
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Apply the display and lookup properties (Caption, Format, ColumnWidth,
'           : DisplayControl, RowSource and the rest) through the same SetDAOProperty
'           : path used when relinking a table.
'---------------------------------------------------------------------------------------
'
Private Sub ApplyFieldProperties(fld As DAO.Field, ByVal dProps As Dictionary)

    Dim varName As Variant

    ' CInt on the type for the same ByRef reason: the dictionary hands back a Variant, and
    ' SetDAOProperty declares an Integer.
    For Each varName In dProps.Keys
        If Not IsNativeFieldMember(CStr(varName)) Then
            SetDAOProperty fld, CInt(dProps(varName)("Type")), CStr(varName), _
                dProps(varName)("Value")
        End If
    Next varName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ApplyTableProperties
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Apply table level properties. ValidationRule and ValidationText are native
'           : TableDef members; everything else is a custom property.
'---------------------------------------------------------------------------------------
'
Private Sub ApplyTableProperties(tdf As DAO.TableDef, ByVal dProps As Dictionary)

    Dim varName As Variant

    If dProps.Exists("ValidationRule") Then tdf.ValidationRule = dProps("ValidationRule")("Value")
    If dProps.Exists("ValidationText") Then tdf.ValidationText = dProps("ValidationText")("Value")

    For Each varName In dProps.Keys
        Select Case CStr(varName)
            Case "ValidationRule", "ValidationText"
                ' Already applied as native members
            Case Else
                SetDAOProperty tdf, CInt(dProps(varName)("Type")), CStr(varName), _
                    dProps(varName)("Value")
        End Select
    Next varName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsNativeFieldMember
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Returns true for field properties handled by ApplyNativeFieldMembers.
'---------------------------------------------------------------------------------------
'
Private Function IsNativeFieldMember(strName As String) As Boolean
    Select Case strName
        Case "Required", "AllowZeroLength", "DefaultValue", "ValidationRule", "ValidationText"
            IsNativeFieldMember = True
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : AppendIndex
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Create one index on the unsaved TableDef.
'---------------------------------------------------------------------------------------
'
Private Sub AppendIndex(tdf As DAO.TableDef, ByVal dIndex As Dictionary)

    Dim idx As DAO.Index
    Dim fld As DAO.Field
    Dim varField As Variant

    Set idx = tdf.CreateIndex(CStr(dIndex("Name")))

    For Each varField In dIndex("Fields")
        Set fld = idx.CreateField(CStr(varField("Name")))
        If varField("Descending") Then fld.Attributes = fld.Attributes Or dbDescending
        idx.Fields.Append fld
    Next varField

    ' Unique first: DAO rejects Primary on an index that is not yet unique.
    idx.Unique = dIndex("Unique")
    idx.Primary = dIndex("Primary")

    tdf.Indexes.Append idx

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemovePartialTable
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Drop a table left behind by a failed build, so the caller's fallback to
'           : Application.ImportXML lands on the intended name.
'---------------------------------------------------------------------------------------
'
Private Sub RemovePartialTable(strTableName As String)

    LogUnhandledErrors
    On Error Resume Next

    DoCmd.DeleteObject acTable, strTableName

    ' If this fails the fallback import lands on an adjusted name rather than raising,
    ' so it is worth a warning even though the operation continues.
    CatchAny eelWarning, "Unable to remove partially built table '" & strTableName & "'", _
        ModuleName & ".RemovePartialTable"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Small parsing helpers
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Attribute and node lookups used throughout the parse.
'---------------------------------------------------------------------------------------
'
Private Function GetAttr(ByVal objNode As MSXML2.IXMLDOMNode, ByVal strName As String) As String
    Dim objAttr As MSXML2.IXMLDOMNode
    If objNode Is Nothing Then Exit Function
    Set objAttr = objNode.Attributes.getNamedItem(strName)
    If Not objAttr Is Nothing Then GetAttr = objAttr.Text
End Function


Private Function GetAttrNs(ByVal objNode As MSXML2.IXMLDOMNode, ByVal strNamespace As String, _
    ByVal strName As String) As String
    Dim objAttr As MSXML2.IXMLDOMNode
    If objNode Is Nothing Then Exit Function
    Set objAttr = objNode.Attributes.getQualifiedItem(strName, strNamespace)
    If Not objAttr Is Nothing Then GetAttrNs = objAttr.Text
End Function


'---------------------------------------------------------------------------------------
' Procedure : FindDescendant
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Return the first descendant with this namespace and local name. Used rather
'           : than XPath because the documents carry no fixed namespace prefixes we can
'           : bind a selection namespace to.
'---------------------------------------------------------------------------------------
'
Private Function FindDescendant(ByVal objNode As MSXML2.IXMLDOMNode, ByVal strNamespace As String, _
    ByVal strLocalName As String) As MSXML2.IXMLDOMNode

    Dim objChild As MSXML2.IXMLDOMNode
    Dim objFound As MSXML2.IXMLDOMNode

    If objNode Is Nothing Then Exit Function

    For Each objChild In objNode.ChildNodes
        If objChild.NodeType = NODE_ELEMENT Then
            If objChild.NamespaceURI = strNamespace And objChild.baseName = strLocalName Then
                Set FindDescendant = objChild
                Exit Function
            End If
            Set objFound = FindDescendant(objChild, strNamespace, strLocalName)
            If Not objFound Is Nothing Then
                Set FindDescendant = objFound
                Exit Function
            End If
        End If
    Next objChild

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFacetValue
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Read the value of an xsd facet (maxLength, totalDigits, fractionDigits)
'           : from a field's simpleType restriction. Returns 0 when absent.
'---------------------------------------------------------------------------------------
'
Private Function GetFacetValue(ByVal objField As MSXML2.IXMLDOMNode, ByVal strFacet As String) As Long
    Dim objFacet As MSXML2.IXMLDOMNode
    Set objFacet = FindDescendant(objField, NS_XSD, strFacet)
    If Not objFacet Is Nothing Then GetFacetValue = CLng(Val(GetAttr(objFacet, "value")))
End Function


Private Function BuildPropertyEntry(intType As Integer, varValue As Variant) As Dictionary
    Set BuildPropertyEntry = New Dictionary
    BuildPropertyEntry.Add "Type", intType
    BuildPropertyEntry.Add "Value", varValue
End Function


Private Function IsYes(strValue As String) As Boolean
    IsYes = (StrComp(strValue, "yes", vbTextCompare) = 0)
End Function


Private Function IsDescending(varOrders As Variant, lngIndex As Long) As Boolean
    If Not IsArray(varOrders) Then Exit Function
    If lngIndex > UBound(varOrders) Then Exit Function
    IsDescending = (StrComp(CStr(varOrders(lngIndex)), "desc", vbTextCompare) = 0)
End Function
