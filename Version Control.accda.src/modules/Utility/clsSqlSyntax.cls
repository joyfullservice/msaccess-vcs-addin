VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSqlSyntax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsSqlSyntax
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Pure SQL token/identifier parsing helpers for clsQueryComposer.
'           : Functional core: no Log, Options, or mutable composer state.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Utility")


'---------------------------------------------------------------------------------------
' Procedure : SplitTableRef
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Parse a FROM-clause table operand into name and optional alias.
'           : Bracket-aware AS detection; no naive space splitting.
'---------------------------------------------------------------------------------------
'
Public Sub SplitTableRef(ByVal strTableExpr As String, _
                         ByRef strName As String, _
                         ByRef strAlias As String)
    Dim strWork As String
    Dim lngAs As Long

    strWork = Trim$(strTableExpr)
    lngAs = FindTopLevelKeyword(strWork, " AS ")
    If lngAs > 0 Then
        strAlias = Trim$(Mid$(strWork, lngAs + 4))
        strName = Trim$(Left$(strWork, lngAs - 1))
    Else
        strAlias = vbNullString
        strName = strWork
    End If

    strName = UnBracket(strName)
    If Len(strAlias) > 0 Then strAlias = UnBracket(strAlias)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : TableRefKey
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Return the join-graph key for a parsed table ref (alias when present).
'---------------------------------------------------------------------------------------
'
Public Function TableRefKey(ByVal strName As String, ByVal strAlias As String) As String
    If Len(strAlias) > 0 Then
        TableRefKey = strAlias
    Else
        TableRefKey = strName
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : TryExtractSimpleTable
' Author    : Adam Waller
' Date      : 4/27/2026
' Purpose   : Returns the table reference key if strTableExpr is a simple table
'           : reference (optionally with an alias), or vbNullString when it
'           : contains a nested JOIN expression.
'           :
'           : EDGE CASE: Bracketed names with spaces (e.g. [Car Models])
'           : must never be split on interior spaces. SplitTableRef and
'           : FindTopLevelKeyword keep name and alias parsing aligned with
'           : AddInputTable so join operands and InputTables keys cannot diverge.
'---------------------------------------------------------------------------------------
'
Public Function TryExtractSimpleTable(strTableExpr As String) As String
    Dim strClean As String
    Dim strName As String
    Dim strAlias As String

    strClean = Trim$(strTableExpr)

    Do While Left$(strClean, 1) = "(" And Right$(strClean, 1) = ")"
        If FindMatchingParen(strClean, 1) = Len(strClean) Then
            strClean = Trim$(Mid$(strClean, 2, Len(strClean) - 2))
        Else
            Exit Do
        End If
    Loop

    If FindTopLevelKeyword(strClean, " JOIN ") > 0 Then
        TryExtractSimpleTable = vbNullString
        Exit Function
    End If

    If Left$(strClean, 1) = "(" Then
        SplitTableRef strTableExpr, strName, strAlias
        TryExtractSimpleTable = TableRefKey(strName, strAlias)
        Exit Function
    End If

    SplitTableRef strClean, strName, strAlias
    TryExtractSimpleTable = TableRefKey(strName, strAlias)
End Function


'---------------------------------------------------------------------------------------
' Procedure : OnClauseTableExcluding
' Author    : Adam Waller
' Date      : 4/27/2026
' Purpose   : Extract both table names from the first '=' in an ON expression and
'           : return whichever is NOT strExcludeTable.
'---------------------------------------------------------------------------------------
'
Public Function OnClauseTableExcluding(strOnExpr As String, _
                                       strExcludeTable As String) As String
    Dim strTblA As String
    Dim strTblB As String
    strTblA = ExtractTableFromOnSide(strOnExpr, True)
    strTblB = ExtractTableFromOnSide(strOnExpr, False)

    If StrComp(strTblA, strExcludeTable, vbTextCompare) = 0 Then
        OnClauseTableExcluding = strTblB
    ElseIf StrComp(strTblB, strExcludeTable, vbTextCompare) = 0 Then
        OnClauseTableExcluding = strTblA
    ElseIf Len(strTblA) > 0 Then
        OnClauseTableExcluding = strTblA
    Else
        OnClauseTableExcluding = strTblB
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExtractTableFromOnSide
' Author    : Adam Waller
' Date      : 4/27/2026
' Purpose   : Extract a table name from one side of the first '=' in an ON expression.
'           : blnLeftSide=True returns the table before '=', False returns after '='.
'           :
'           : Returns empty when the side is not a single bare or qualified identifier
'           : (e.g. DateAdd('yyyy', -1, cur.EffectiveDate) or Left(a.x, 3)). Callers
'           : must fall back to a known InputTables ref rather than trusting a
'           : non-empty garbage token. See docs/access-query-storage.md.
'---------------------------------------------------------------------------------------
'
Public Function ExtractTableFromOnSide(strOnExpr As String, blnLeftSide As Boolean) As String
    Dim lngEq As Long
    Dim strSide As String
    Dim lngDot As Long

    lngEq = InStr(1, strOnExpr, "=")
    If lngEq = 0 Then Exit Function

    If blnLeftSide Then
        strSide = Trim$(Left$(strOnExpr, lngEq - 1))
    Else
        strSide = Trim$(Mid$(strOnExpr, lngEq + 1))
    End If

    Do While Left$(strSide, 1) = "("
        strSide = Mid$(strSide, 2)
    Loop
    Do While Right$(strSide, 1) = ")"
        strSide = Left$(strSide, Len(strSide) - 1)
    Loop
    strSide = Trim$(strSide)

    ' Reject function calls, operators, and multi-token expressions so we never
    ' return a prefix like "DateAdd('yyyy', -1, cur" as a table name.
    If Not IsSimpleOnSideIdentifier(strSide) Then Exit Function

    lngDot = FindQualifyingDot(strSide)
    If lngDot > 0 Then
        ExtractTableFromOnSide = UnBracket(Trim$(Left$(strSide, lngDot - 1)))
    Else
        ExtractTableFromOnSide = UnBracket(strSide)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsSimpleOnSideIdentifier
' Author    : Adam Waller
' Date      : 8/10/2026
' Purpose   : True when strSide is a bare or qualified identifier suitable for
'           : ExtractTableFromOnSide — optionally bracketed parts separated by a
'           : single qualifying dot. False for function calls, operators, commas,
'           : or spaces outside brackets/quotes.
'---------------------------------------------------------------------------------------
'
Private Function IsSimpleOnSideIdentifier(ByVal strSide As String) As Boolean
    Dim lngPos As Long
    Dim ch As String
    Dim blnInBracket As Boolean
    Dim blnInString As Boolean

    If Len(strSide) = 0 Then Exit Function

    For lngPos = 1 To Len(strSide)
        ch = Mid$(strSide, lngPos, 1)
        If blnInString Then
            If ch = "'" Then blnInString = False
        ElseIf blnInBracket Then
            If ch = "]" Then blnInBracket = False
        Else
            Select Case ch
                Case "["
                    blnInBracket = True
                Case "'"
                    blnInString = True
                Case "(", ")", ",", "+", "-", "*", "/", "&", "<", ">", "!", " "
                    Exit Function
            End Select
        End If
    Next lngPos

    ' Unclosed bracket/string is not a simple identifier.
    If blnInBracket Or blnInString Then Exit Function
    IsSimpleOnSideIdentifier = True
End Function


Private Function FindQualifyingDot(ByVal strSide As String) As Long
    Dim lngPos As Long
    Dim blnInBracket As Boolean

    For lngPos = 1 To Len(strSide)
        Select Case Mid$(strSide, lngPos, 1)
            Case "["
                blnInBracket = True
            Case "]"
                blnInBracket = False
            Case "."
                If Not blnInBracket Then
                    FindQualifyingDot = lngPos
                    Exit Function
                End If
        End Select
    Next lngPos
End Function


Public Function FindTopLevelKeyword(strSql As String, strKeyword As String) As Long
    Dim lngPos As Long
    Dim lngDepth As Long
    Dim blnInString As Boolean
    Dim blnInBracket As Boolean
    Dim strUpper As String
    Dim strKeyUpper As String

    strUpper = UCase$(strSql)
    strKeyUpper = UCase$(strKeyword)

    For lngPos = 1 To Len(strUpper) - Len(strKeyUpper) + 1
        Select Case Mid$(strUpper, lngPos, 1)
            Case """"
                If Not blnInBracket Then blnInString = Not blnInString
            Case "["
                If Not blnInString Then blnInBracket = True
            Case "]"
                If Not blnInString Then blnInBracket = False
            Case "("
                If Not blnInString And Not blnInBracket Then lngDepth = lngDepth + 1
            Case ")"
                If Not blnInString And Not blnInBracket Then lngDepth = lngDepth - 1
        End Select

        If Not blnInString And Not blnInBracket And lngDepth = 0 Then
            If Mid$(strUpper, lngPos, Len(strKeyUpper)) = strKeyUpper Then
                If Left$(strKeyUpper, 1) = " " Then
                    FindTopLevelKeyword = lngPos
                    Exit Function
                ElseIf lngPos = 1 Then
                    FindTopLevelKeyword = lngPos
                    Exit Function
                ElseIf Mid$(strUpper, lngPos - 1, 1) = " " Or _
                   Mid$(strUpper, lngPos - 1, 1) = ")" Then
                    FindTopLevelKeyword = lngPos
                    Exit Function
                End If
            End If
        End If
    Next lngPos
End Function


Public Function FindJoinKeyword(strExpr As String, ByRef strJoinType As String, ByRef lngFlag As Long) As Long
    Dim lngPos As Long

    lngPos = FindTopLevelKeyword(strExpr, " INNER JOIN ")
    If lngPos > 0 Then
        strJoinType = " INNER JOIN "
        lngFlag = 1
        FindJoinKeyword = lngPos
        Exit Function
    End If

    lngPos = FindTopLevelKeyword(strExpr, " LEFT JOIN ")
    If lngPos > 0 Then
        strJoinType = " LEFT JOIN "
        lngFlag = 2
        FindJoinKeyword = lngPos
        Exit Function
    End If

    lngPos = FindTopLevelKeyword(strExpr, " RIGHT JOIN ")
    If lngPos > 0 Then
        strJoinType = " RIGHT JOIN "
        lngFlag = 3
        FindJoinKeyword = lngPos
        Exit Function
    End If
End Function


Public Function FindLastJoinKeyword(strExpr As String, ByRef strJoinType As String, ByRef lngFlag As Long) As Long
    Dim lngPos As Long
    Dim lngLast As Long
    Dim strLastType As String
    Dim lngLastFlag As Long
    Dim lngSearch As Long
    Dim strTmpType As String
    Dim lngTmpFlag As Long
    Dim lngInner As Long
    Dim lngLeft As Long
    Dim lngRight As Long

    lngSearch = 1
    lngLast = 0

    Do
        lngPos = 0
        lngInner = FindTopLevelKeywordFrom(strExpr, " INNER JOIN ", lngSearch)
        lngLeft = FindTopLevelKeywordFrom(strExpr, " LEFT JOIN ", lngSearch)
        lngRight = FindTopLevelKeywordFrom(strExpr, " RIGHT JOIN ", lngSearch)

        If lngInner > 0 Then
            lngPos = lngInner: strTmpType = " INNER JOIN ": lngTmpFlag = 1
        End If
        If lngLeft > 0 And (lngPos = 0 Or lngLeft < lngPos) Then
            lngPos = lngLeft: strTmpType = " LEFT JOIN ": lngTmpFlag = 2
        End If
        If lngRight > 0 And (lngPos = 0 Or lngRight < lngPos) Then
            lngPos = lngRight: strTmpType = " RIGHT JOIN ": lngTmpFlag = 3
        End If

        If lngPos > 0 Then
            lngLast = lngPos
            strLastType = strTmpType
            lngLastFlag = lngTmpFlag
            lngSearch = lngPos + Len(strTmpType)
        End If
    Loop While lngPos > 0

    If lngLast > 0 Then
        FindLastJoinKeyword = lngLast
        strJoinType = strLastType
        lngFlag = lngLastFlag
    End If
End Function


Public Function FindTopLevelKeywordFrom(strSql As String, strKeyword As String, lngStart As Long) As Long
    Dim lngPos As Long
    Dim lngDepth As Long
    Dim blnInString As Boolean
    Dim strUpper As String
    Dim strKeyUpper As String

    strUpper = UCase$(strSql)
    strKeyUpper = UCase$(strKeyword)

    For lngPos = lngStart To Len(strUpper) - Len(strKeyUpper) + 1
        Select Case Mid$(strUpper, lngPos, 1)
            Case """"
                blnInString = Not blnInString
            Case "("
                If Not blnInString Then lngDepth = lngDepth + 1
            Case ")"
                If Not blnInString Then lngDepth = lngDepth - 1
        End Select

        If Not blnInString And lngDepth = 0 Then
            If Mid$(strUpper, lngPos, Len(strKeyUpper)) = strKeyUpper Then
                FindTopLevelKeywordFrom = lngPos
                Exit Function
            End If
        End If
    Next lngPos
End Function


Public Function FindMatchingParen(strExpr As String, lngStart As Long) As Long
    Dim lngPos As Long
    Dim lngDepth As Long
    Dim blnInString As Boolean

    lngDepth = 0
    For lngPos = lngStart To Len(strExpr)
        Select Case Mid$(strExpr, lngPos, 1)
            Case """"
                blnInString = Not blnInString
            Case "("
                If Not blnInString Then lngDepth = lngDepth + 1
            Case ")"
                If Not blnInString Then
                    lngDepth = lngDepth - 1
                    If lngDepth = 0 Then
                        FindMatchingParen = lngPos
                        Exit Function
                    End If
                End If
        End Select
    Next lngPos
End Function


Public Function HasTopLevelBoolean(strExpr As String) As Boolean
    Dim strWork As String

    strWork = Trim$(strExpr)
    If Len(strWork) = 0 Then Exit Function

    Do While Left$(strWork, 1) = "(" And Right$(strWork, 1) = ")"
        If FindMatchingParen(strWork, 1) = Len(strWork) Then
            strWork = Trim$(Mid$(strWork, 2, Len(strWork) - 2))
        Else
            Exit Do
        End If
    Loop

    HasTopLevelBoolean = (FindTopLevelKeyword(strWork, " AND ") > 0) Or _
                         (FindTopLevelKeyword(strWork, " OR ") > 0)
End Function


Public Function StripOuterParens(strExpr As String) As String
    Dim strWork As String
    strWork = Trim$(strExpr)
    Do While Left$(strWork, 1) = "(" And Right$(strWork, 1) = ")"
        If FindMatchingParen(strWork, 1) = Len(strWork) Then
            strWork = Trim$(Mid$(strWork, 2, Len(strWork) - 2))
        Else
            Exit Do
        End If
    Loop
    StripOuterParens = strWork
End Function


Public Function SplitTopLevel(strExpr As String, strDelim As String) As String()
    Dim colParts As New Collection
    Dim lngStart As Long
    Dim lngPos As Long
    Dim lngDepth As Long
    Dim blnInDblQuote As Boolean
    Dim blnInSglQuote As Boolean
    Dim blnInBracket As Boolean
    Dim asParts() As String
    Dim blnInAnyString As Boolean
    Dim lngI As Long

    lngStart = 1
    For lngPos = 1 To Len(strExpr)
        blnInAnyString = blnInDblQuote Or blnInSglQuote
        Select Case Mid$(strExpr, lngPos, 1)
            Case """"
                If Not blnInBracket And Not blnInSglQuote Then blnInDblQuote = Not blnInDblQuote
            Case "'"
                If Not blnInBracket And Not blnInDblQuote Then blnInSglQuote = Not blnInSglQuote
            Case "["
                If Not blnInAnyString Then blnInBracket = True
            Case "]"
                If Not blnInAnyString Then blnInBracket = False
            Case "("
                If Not blnInAnyString And Not blnInBracket Then lngDepth = lngDepth + 1
            Case ")"
                If Not blnInAnyString And Not blnInBracket Then lngDepth = lngDepth - 1
            Case Else
                If Mid$(strExpr, lngPos, Len(strDelim)) = strDelim Then
                    If Not blnInAnyString And Not blnInBracket And lngDepth = 0 Then
                        colParts.Add Mid$(strExpr, lngStart, lngPos - lngStart)
                        lngStart = lngPos + Len(strDelim)
                    End If
                End If
        End Select
    Next lngPos

    colParts.Add Mid$(strExpr, lngStart)

    ReDim asParts(0 To colParts.Count - 1)
    For lngI = 1 To colParts.Count
        asParts(lngI - 1) = CStr(colParts(lngI))
    Next lngI

    SplitTopLevel = asParts
End Function


Public Function UnBracket(strName As String) As String
    strName = Trim$(strName)
    If Left$(strName, 1) = "[" And Right$(strName, 1) = "]" Then
        UnBracket = Mid$(strName, 2, Len(strName) - 2)
    Else
        UnBracket = strName
    End If
End Function


Public Function BracketIfNeeded(strName As String) As String
    If Left$(strName, 1) = "[" Or Left$(strName, 1) = "(" Then
        BracketIfNeeded = strName
    ElseIf InStr(1, strName, ".") > 0 Then
        BracketIfNeeded = strName
    ElseIf HasNonIdentChars(strName) Or IsAccessReservedWord(strName) Then
        BracketIfNeeded = "[" & strName & "]"
    Else
        BracketIfNeeded = strName
    End If
End Function


Public Function HasNonIdentChars(strName As String) As Boolean
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(strName)
        ch = Mid$(strName, i, 1)
        Select Case ch
            Case "A" To "Z", "a" To "z", "0" To "9", "_"
            Case Else
                HasNonIdentChars = True
                Exit Function
        End Select
    Next i
End Function


Public Function IsAccessReservedWord(strWord As String) As Boolean
    Select Case UCase$(strWord)
        Case "ACTION", "SELECT", "FROM", "WHERE", "ORDER", "BY", "GROUP", "HAVING", _
             "INSERT", "INTO", "UPDATE", "DELETE", "SET", "VALUE", "VALUES", _
             "JOIN", "INNER", "LEFT", "RIGHT", "OUTER", "ON", "AS", _
             "AND", "OR", "NOT", "IN", "BETWEEN", "LIKE", "IS", "NULL", _
             "EXISTS", "ALL", "ANY", "SOME", "UNION", "EXCEPT", "INTERSECT", _
             "DISTINCT", "TOP", "TRANSFORM", "PIVOT", "TABLE", "CREATE", _
             "ALTER", "DROP", "INDEX", "VIEW", "PROCEDURE", "PARAMETER", _
             "PARAMETERS", "TRUE", "FALSE", "YES", "NO", "DATE", "TIME", "YEAR", _
             "TEXT", "MEMO", "LONG", "SHORT", "SINGLE", "DOUBLE", "CURRENCY", _
             "COUNTER", "YESNO", "BINARY", "OLEOBJECT", "GUID", "KEY", "NAMES", "NUMBER", _
             "PRIMARY", "FOREIGN", "CONSTRAINT", "REFERENCES", "UNIQUE"
            IsAccessReservedWord = True
    End Select
End Function


Public Function IsIdentChar(strCh As String) As Boolean
    Dim lngCode As Long
    lngCode = AscW(strCh)
    IsIdentChar = (lngCode >= 65 And lngCode <= 90) Or _
                  (lngCode >= 97 And lngCode <= 122) Or _
                  (lngCode >= 48 And lngCode <= 57) Or _
                  lngCode = 95 Or lngCode = 91 Or lngCode = 93
End Function
