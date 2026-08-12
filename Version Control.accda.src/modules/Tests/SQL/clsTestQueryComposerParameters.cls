VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTestQueryComposerParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsTestQueryComposerParameters
' Author    : Adam Waller
' Date      : 8/11/2026
' Purpose   : Declared query parameter tests for clsQueryComposer.
'           : The expectations here were captured from Access rather than inferred from
'           : the .qdef grammar: parameterized queries of each shape were saved through
'           : the query designer and re-exported with Application.SaveAsText, and each
'           : type keyword was round-tripped through CreateQueryDef to learn the DAO type
'           : it actually produces. See DECISIONS.md and docs/access-query-storage.md.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Tests.SQL")
'@Tag("unit")

' Type keywords the ACE parser accepts in a PARAMETERS clause. Spellings absent from this
' list -- notably "Boolean", "Memo", "OLEObject" and "Decimal" -- raise error 3139, so
' they may be accepted on import but must never be emitted.
Private Const ACCEPTED_KEYWORDS As String = _
    "|value|bit|yesno|logical|byte|short|smallint|integer|int|long|currency|money|" & _
    "single|ieeesingle|real|double|ieeedouble|float|number|numeric|datetime|date|time|" & _
    "binary|varbinary|text|char|varchar|string|longtext|note|longbinary|guid|bigint|"

Private m_qc As clsQueryComposer


Private Sub Class_Initialize()
    Set m_qc = New clsQueryComposer
End Sub


Private Sub Class_Terminate()
    Set m_qc = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestParameterBlockFollowsOutputColumns
' Purpose   : Access accepts "Begin Parameters" in exactly one place -- immediately after
'           : the OutputColumns block. Every other position was rejected by LoadFromText
'           : ("Expected: End of file. Found: Parameters."), so this is a hard constraint
'           : rather than a stylistic choice.
'---------------------------------------------------------------------------------------
'
Public Sub TestParameterBlockFollowsOutputColumns()

    Dim strQdef As String

    strQdef = DesignQdef("PARAMETERS [Min Qty] Long; " & _
        "SELECT t1.Category, t2.Label FROM t1 INNER JOIN t2 ON t1.ID = t2.A_ID " & _
        "WHERE (((t1.Qty)>=[Min Qty])) ORDER BY t1.Category")

    TestAssert InStr(strQdef, "Begin Parameters") > 0, "parameters block emitted"
    TestAssert BlockOrder(strQdef, "Begin OutputColumns", "Begin Parameters"), _
        "parameters block follows output columns"
    TestAssert BlockOrder(strQdef, "Begin Parameters", "Begin Joins"), _
        "parameters block precedes joins"
    TestAssert BlockOrder(strQdef, "Begin Parameters", "Begin OrderBy"), _
        "parameters block precedes order by"

End Sub


Public Sub TestParameterBlockPrecedesGroups()

    Dim strQdef As String

    strQdef = DesignQdef("PARAMETERS [Min Qty] Long; " & _
        "SELECT t1.Category FROM t1 WHERE (((t1.Qty)>=[Min Qty])) GROUP BY t1.Category")

    TestAssert BlockOrder(strQdef, "Begin OutputColumns", "Begin Parameters"), _
        "parameters block follows output columns"
    TestAssert BlockOrder(strQdef, "Begin Parameters", "Begin Groups"), _
        "parameters block precedes groups"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestParameterNamesKeptVerbatim
' Purpose   : Access records a parameter name exactly as it was declared. A bracketed
'           : name keeps its brackets and an unbracketed one stays bare, even though
'           : Access brackets the matching reference inside the WHERE clause.
'---------------------------------------------------------------------------------------
'
Public Sub TestParameterNamesKeptVerbatim()

    Dim strQdef As String

    strQdef = DesignQdef("PARAMETERS [Min Qty] Long, StatusFilter Short; " & _
        "SELECT t1.Category FROM t1 " & _
        "WHERE (((t1.Qty)>=[Min Qty]) AND ((t1.StatusID)=StatusFilter))")

    TestAssert InStr(strQdef, "Name =""[Min Qty]""") > 0, _
        "bracketed parameter name keeps its brackets"
    TestAssert InStr(strQdef, "Name =""StatusFilter""") > 0, _
        "unbracketed parameter name is not bracketed"

End Sub


Public Sub TestMultipleParametersShareOneBlock()

    Dim strQdef As String
    Dim strBlock As String

    strQdef = DesignQdef("PARAMETERS [Enter ID] Long, [Enter Category] Text ( 255 ); " & _
        "SELECT t1.ID FROM t1 " & _
        "WHERE (((t1.ID)=[Enter ID]) AND ((t1.Category)=[Enter Category]))")
    strBlock = ParameterBlock(strQdef)

    TestAssert UBound(Split(strQdef, "Begin Parameters")) = 1, _
        "all parameters share a single block"
    TestAssert ParameterCount(strQdef) = 2, "both parameters emitted"
    TestAssert InStr(strBlock, "Flag =4" & vbCrLf) > 0, "first parameter typed as Long"
    TestAssert InStr(strBlock, "Flag =10" & vbCrLf) > 0, "second parameter typed as Text"

End Sub


Public Sub TestQueryWithoutParametersEmitsNoBlock()

    Dim strQdef As String

    strQdef = DesignQdef("SELECT t1.ID FROM t1")
    TestAssert InStr(strQdef, "Begin Parameters") = 0, _
        "no parameters block without a PARAMETERS clause"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestBracketedNameContainingComma
' Purpose   : A bracketed parameter name may legally contain a comma, which must not be
'           : treated as a separator between declarations.
'---------------------------------------------------------------------------------------
'
Public Sub TestBracketedNameContainingComma()

    Dim strQdef As String

    strQdef = DesignQdef("PARAMETERS [Last, First] Text ( 255 ); " & _
        "SELECT t1.ID FROM t1 WHERE (((t1.FullName)=[Last, First]))")

    TestAssert ParameterCount(strQdef) = 1, "comma inside brackets is not a separator"
    TestAssert InStr(strQdef, "Name =""[Last, First]""") > 0, "name preserved intact"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestParameterTypeFlags
' Purpose   : Each keyword maps to the DAO data type Access itself assigns. The aliases
'           : below were confirmed against the ACE parser; "Boolean", "Memo" and
'           : "OLEObject" are import-only spellings retained so source exported by earlier
'           : versions of the add-in still resolves to the right type.
'---------------------------------------------------------------------------------------
'
Public Sub TestParameterTypeFlags()

    AssertParameterFlag "Bit", dbBoolean
    AssertParameterFlag "YesNo", dbBoolean
    AssertParameterFlag "Boolean", dbBoolean
    AssertParameterFlag "Byte", dbByte
    AssertParameterFlag "Short", dbInteger
    AssertParameterFlag "SmallInt", dbInteger
    AssertParameterFlag "Long", dbLong
    AssertParameterFlag "Integer", dbLong
    AssertParameterFlag "Currency", dbCurrency
    AssertParameterFlag "Single", dbSingle
    AssertParameterFlag "IEEESingle", dbSingle
    AssertParameterFlag "Double", dbDouble
    AssertParameterFlag "IEEEDouble", dbDouble
    AssertParameterFlag "Numeric", dbDouble
    AssertParameterFlag "DateTime", dbDate
    AssertParameterFlag "Binary", dbBinary
    AssertParameterFlag "Text", dbText
    AssertParameterFlag "Value", dbText
    AssertParameterFlag "LongBinary", dbLongBinary
    AssertParameterFlag "OLEObject", dbLongBinary
    AssertParameterFlag "LongText", dbMemo
    AssertParameterFlag "Memo", dbMemo
    AssertParameterFlag "GUID", dbGUID
    AssertParameterFlag "BigInt", dbNumeric

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestSizeSpecifierIgnored
' Purpose   : A size specifier belongs to the SQL PARAMETERS clause only; the Design View
'           : block carries the bare DAO type. Spacing around it varies by author.
'---------------------------------------------------------------------------------------
'
Public Sub TestSizeSpecifierIgnored()

    AssertParameterFlag "Text ( 255 )", dbText
    AssertParameterFlag "Text(50)", dbText
    AssertParameterFlag "VarChar ( 12 )", dbText

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestExportedTypeKeywordsAreParseable
' Purpose   : Close the loop on the type table. For every DAO type Access can store, the
'           : export path must emit a keyword the ACE parser accepts, and re-importing
'           : that keyword must yield the same type. This is what catches a regression to
'           : an unparseable spelling such as "Boolean".
'---------------------------------------------------------------------------------------
'
Public Sub TestExportedTypeKeywordsAreParseable()

    Dim vFlags As Variant
    Dim i As Long
    Dim lngFlag As Long
    Dim strSql As String
    Dim strKeyword As String

    vFlags = Array(dbBoolean, dbByte, dbInteger, dbLong, dbCurrency, dbSingle, _
        dbDouble, dbDate, dbBinary, dbText, dbLongBinary, dbMemo, dbGUID, dbNumeric)

    For i = LBound(vFlags) To UBound(vFlags)
        lngFlag = CLng(vFlags(i))
        strSql = ReconstructParameterSql(lngFlag)
        strKeyword = ParameterKeywordInSql(strSql)
        TestAssert IsAcceptedKeyword(strKeyword), _
            "DAO type " & lngFlag & " emits a parseable keyword ('" & strKeyword & "')"
        TestAssert ParameterFlagInQdef(DesignQdef(strSql)) = lngFlag, _
            "DAO type " & lngFlag & " survives the export/import round trip"
    Next i

End Sub


Public Sub TestUnrecognizedTypeWarns()

    Dim strQdef As String
    Dim eimPriorMode As eInteractionMode

    ' This is the only test here that provokes a real builder warning, and clsLog
    ' turns a warning into a modal MsgBox whenever no console is attached to display
    ' it. Whether one is attached depends on how the suite was launched, so force
    ' silent mode for the duration rather than inherit the caller's mode -- the same
    ' guard modTestRoundtrip uses to keep prompts from stalling an unattended run.
    eimPriorMode = Operation.InteractionMode
    Operation.InteractionMode = eimSilent
    Err.Clear
    On Error GoTo RestoreMode

    strQdef = DesignQdef("PARAMETERS [P] Frobnicate; " & _
        "SELECT t1.ID FROM t1 WHERE (((t1.ID)=[P]))")

RestoreMode:
    Operation.InteractionMode = eimPriorMode
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
    On Error GoTo 0

    TestAssert m_qc.WarningCount > 0, "unrecognized parameter type raises a warning"
    TestAssert ParameterFlagInQdef(strQdef) = 0, _
        "unrecognized parameter type falls back to an untyped flag"

End Sub


'=======================================================================================
' PRIVATE: Helpers
'=======================================================================================

'---------------------------------------------------------------------------------------
' Procedure : DesignQdef
' Purpose   : Decompose SQL and render it as a Design View .qdef, leaving m_qc pointing at
'           : the composer used so the caller can inspect its warnings.
'---------------------------------------------------------------------------------------
'
Private Function DesignQdef(ByVal strSql As String) As String
    Set m_qc = New clsQueryComposer
    m_qc.ContextLabel = "clsTestQueryComposerParameters"
    m_qc.DecomposeSQL strSql
    DesignQdef = m_qc.GenerateQdef(True, New Dictionary, New Dictionary)
End Function


Private Sub AssertParameterFlag(ByVal strKeyword As String, ByVal lngExpected As Long)
    Dim strQdef As String
    strQdef = DesignQdef("PARAMETERS [P] " & strKeyword & "; " & _
        "SELECT t1.ID FROM t1 WHERE (((t1.ID)=[P]))")
    TestAssert ParameterFlagInQdef(strQdef) = lngExpected, _
        "'" & strKeyword & "' maps to DAO type " & lngExpected
End Sub


Private Function BlockOrder(ByVal strQdef As String, ByVal strFirst As String, _
    ByVal strSecond As String) As Boolean

    Dim lngFirst As Long
    Dim lngSecond As Long

    lngFirst = InStr(1, strQdef, strFirst)
    lngSecond = InStr(1, strQdef, strSecond)
    BlockOrder = (lngFirst > 0) And (lngSecond > 0) And (lngFirst < lngSecond)

End Function


Private Function ParameterBlock(ByVal strQdef As String) As String

    Dim lngStart As Long
    Dim lngEnd As Long

    lngStart = InStr(1, strQdef, "Begin Parameters")
    If lngStart = 0 Then Exit Function
    lngEnd = InStr(lngStart, strQdef, vbCrLf & "End" & vbCrLf)
    If lngEnd = 0 Then Exit Function
    ParameterBlock = Mid$(strQdef, lngStart, lngEnd - lngStart + Len(vbCrLf))

End Function


Private Function ParameterCount(ByVal strQdef As String) As Long
    Dim strBlock As String
    strBlock = ParameterBlock(strQdef)
    If Len(strBlock) = 0 Then Exit Function
    ParameterCount = UBound(Split(strBlock, "Name ="))
End Function


Private Function ParameterFlagInQdef(ByVal strQdef As String) As Long

    Dim lngBlock As Long
    Dim lngFlag As Long
    Dim lngEol As Long

    ParameterFlagInQdef = -1
    lngBlock = InStr(1, strQdef, "Begin Parameters")
    If lngBlock = 0 Then Exit Function
    lngFlag = InStr(lngBlock, strQdef, "Flag =")
    If lngFlag = 0 Then Exit Function
    lngEol = InStr(lngFlag, strQdef, vbCrLf)
    If lngEol = 0 Then lngEol = Len(strQdef) + 1
    ParameterFlagInQdef = CLng(Trim$(Mid$(strQdef, lngFlag + 6, lngEol - lngFlag - 6)))

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReconstructParameterSql
' Purpose   : Drive the export path with MSysQueries-shaped rows carrying one declared
'           : parameter of the given DAO type, returning the SQL Access would receive.
'---------------------------------------------------------------------------------------
'
Private Function ReconstructParameterSql(ByVal lngFlag As Long) As String

    Dim colRows As New Collection
    Dim dRow As Dictionary

    Set dRow = New Dictionary
    dRow.Add "Attribute", 1: dRow.Add "Flag", 1
    dRow.Add "Expression", vbNullString: dRow.Add "Name1", vbNullString: dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 2: dRow.Add "Flag", lngFlag
    dRow.Add "Expression", vbNullString: dRow.Add "Name1", "[P]": dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 5: dRow.Add "Flag", 0
    dRow.Add "Expression", vbNullString: dRow.Add "Name1", "t1": dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 6: dRow.Add "Flag", 0
    dRow.Add "Expression", "t1.ID": dRow.Add "Name1", vbNullString: dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set dRow = New Dictionary
    dRow.Add "Attribute", 255: dRow.Add "Flag", 0
    dRow.Add "Expression", vbNullString: dRow.Add "Name1", vbNullString: dRow.Add "Name2", vbNullString
    colRows.Add dRow

    Set m_qc = New clsQueryComposer
    m_qc.ContextLabel = "clsTestQueryComposerParameters"
    ReconstructParameterSql = m_qc.ReconstructSQL(colRows)

End Function


Private Function ParameterKeywordInSql(ByVal strSql As String) As String

    Dim lngStart As Long
    Dim lngEnd As Long

    lngStart = InStr(1, strSql, "[P] ")
    If lngStart = 0 Then Exit Function
    lngStart = lngStart + Len("[P] ")
    lngEnd = InStr(lngStart, strSql, ";")
    If lngEnd = 0 Then Exit Function
    ParameterKeywordInSql = Trim$(Mid$(strSql, lngStart, lngEnd - lngStart))

End Function


Private Function IsAcceptedKeyword(ByVal strKeyword As String) As Boolean

    Dim lngParen As Long

    lngParen = InStr(1, strKeyword, "(")
    If lngParen > 0 Then strKeyword = Left$(strKeyword, lngParen - 1)
    strKeyword = LCase$(Replace$(Trim$(strKeyword), " ", vbNullString))
    IsAcceptedKeyword = (InStr(1, ACCEPTED_KEYWORDS, "|" & strKeyword & "|") > 0)

End Function
