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

' Words the Access query designer brackets when it writes them as an alias, plus
' EQV, IMP, MOD and XOR, which the engine rejects as table aliases. Swept from
' Access itself; see docs/access-query-storage.md. Space-delimited on both ends.
Private Const DESIGNER_RESERVED_WORDS_1 As String = " " & _
    "ABSOLUTE ACTION ADD ADMINDB ALL ALLOCATE ALPHANUMERIC ALTER AND ANY ARE AS ASC " & _
    "ASSERTION AT AUTHORIZATION AUTOINCREMENT AVG BAND BEGIN BETWEEN BINARY BIT " & _
    "BIT_LENGTH BNOT BOR BOTH BXOR BY BYTE CASCADE CASCADED CASE CAST CATALOG CHAR " & _
    "CHAR_LENGTH CHARACTER CHARACTER_LENGTH CHECK CLOSE COALESCE COLLATE COLLATION " & _
    "COLUMN COMMIT COMP COMPRESSION CONNECT CONNECTION CONSTRAINT CONSTRAINTS " & _
    "CONTAINER CONTINUE CONVERT CORRESPONDING COUNT COUNTER CREATE CREATEDB CROSS " & _
    "CURRENCY CURRENT CURRENT_DATE CURRENT_TIME CURRENT_TIMESTAMP CURRENT_USER " & _
    "CURSOR DATABASE DATE DATETIME DAY DEALLOCATE DEC DECIMAL DECLARE DEFAULT " & _
    "DEFERRABLE DEFERRED DELETE DESC DESCRIBE DESCRIPTOR DIAGNOSTICS DISALLOW " & _
    "DISCONNECT DISTINCT DISTINCTROW DOMAIN DOUBLE DROP ELSE END EQV ESCAPE EXCEPT " & _
    "EXCEPTION EXCLUSIVECONNECT EXEC EXECUTE EXISTS EXTERNAL EXTRACT FALSE FETCH " & _
    "FIRST FLOAT FLOAT4 FLOAT8 FOR FOREIGN FOUND FROM FULL GENERAL GET GLOBAL GO " & _
    "GOTO GRANT GROUP GUID HAVING HOUR IDENTITY IEEEDOUBLE IEEESINGLE IGNORE IMAGE " & _
    "IMMEDIATE IMP IN INDEX INDICATOR INHERITABLE INITIALLY INNER INPUT INSENSITIVE " & _
    "INSERT INT INTEGER INTEGER1 INTEGER2 INTEGER4 INTERSECT INTERVAL INTO IS "
Private Const DESIGNER_RESERVED_WORDS_2 As String = _
    "ISOLATION JOIN KEY LANGUAGE LAST LEADING LEFT LEVEL LIKE LOCAL LOGICAL LOGICAL1 " & _
    "LONG LONGBINARY LONGCHAR LONGTEXT LOWER MATCH MAX MEMO MIN MINUTE MOD MODULE " & _
    "MONEY MONTH NAMES NATIONAL NATURAL NCHAR NEXT NO NOT NOTE NULL NULLIF NUMBER " & _
    "NUMERIC OBJECT OCTET_LENGTH OF OLEOBJECT ON ONLY OPEN OPTION OR ORDER OUTER " & _
    "OUTPUT OVERLAPS OWNERACCESS PAD PARAMETERS PARTIAL PASSWORD PERCENT PIVOT " & _
    "POSITION PRECISION PREPARE PRESERVE PRIMARY PRIOR PRIVILEGES PROC PROCEDURE " & _
    "PUBLIC READ REAL REFERENCES RELATIVE RESTRICT REVOKE RIGHT ROLLBACK ROWS SCHEMA " & _
    "SCROLL SECOND SECTION SELECT SELECTSCHEMA SELECTSECURITY SESSION SESSION_USER " & _
    "SET SHORT SINGLE SIZE SMALLINT SOME SPACE SQL SQLCODE SQLERROR SQLSTATE STRING " & _
    "SUBSTRING SUM SYSTEM_USER TABLE TABLEID TEMPORARY TEXT THEN TIME TIMESTAMP " & _
    "TIMEZONE_HOUR TIMEZONE_MINUTE TO TOP TRAILING TRANSACTION TRANSFORM TRANSLATE " & _
    "TRANSLATION TRIM TRUE UNION UNIQUE UNIQUEIDENTIFIER UNKNOWN UPDATE " & _
    "UPDATEIDENTITY UPDATEOWNER UPDATESECURITY UPPER USAGE USER USING VALUE VALUES " & _
    "VARBINARY VARCHAR VARYING VIEW WHEN WHENEVER WHERE WITH WORK WRITE XOR YEAR " & _
    "YESNO ZONE "
Private Const DESIGNER_RESERVED_WORDS As String = DESIGNER_RESERVED_WORDS_1 & DESIGNER_RESERVED_WORDS_2

' Words Access rejects as an unqualified UPDATE SET target (`SET Date = 1`) unless
' bracketed. Qualified (`SET t.Date = 1`) every reserved word parses bare. Swept
' from Access itself; see docs/access-query-storage.md. Space-delimited on both ends.
Private Const SET_TARGET_RESERVED_WORDS As String = " " & _
    "ALL ALPHANUMERIC AS ASC BINARY BIT BY BYTE CHAR CURRENCY DATABASE DATE DATETIME " & _
    "DELETE DESC DISTINCT DISTINCTROW DOUBLE FLOAT FLOAT4 FLOAT8 FROM GENERAL GROUP " & _
    "GUID HAVING IEEEDOUBLE IEEESINGLE IN INNER INSERT INT INTEGER INTEGER1 INTEGER2 " & _
    "INTEGER4 INTO JOIN LEFT LEVEL LOGICAL LOGICAL1 LONG LONGBINARY LONGCHAR LONGTEXT " & _
    "MONEY NOTE NUMBER NUMERIC ON OPTION ORDER OUTER OWNERACCESS PARAMETERS PERCENT " & _
    "PIVOT PROCEDURE REAL RIGHT SELECT SET SHORT SINGLE SMALLINT STRING TABLE TABLEID " & _
    "TEXT TIME TIMESTAMP TOP TRANSFORM UNION UNIQUE UPDATE VALUE VALUES VARBINARY " & _
    "VARCHAR WHERE WITH YESNO "


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
    Dim blnInDouble As Boolean
    Dim blnInSingle As Boolean
    Dim blnInBracket As Boolean
    Dim strUpper As String
    Dim strKeyUpper As String

    strUpper = UCase$(strSql)
    strKeyUpper = UCase$(strKeyword)

    For lngPos = 1 To Len(strUpper) - Len(strKeyUpper) + 1
        Select Case Mid$(strUpper, lngPos, 1)
            Case """"
                If Not blnInSingle And Not blnInBracket Then
                    blnInDouble = Not blnInDouble
                End If
            Case "'"
                If Not blnInDouble And Not blnInBracket Then
                    blnInSingle = Not blnInSingle
                End If
            Case "["
                If Not blnInDouble And Not blnInSingle Then blnInBracket = True
            Case "]"
                If Not blnInDouble And Not blnInSingle Then blnInBracket = False
            Case "("
                If Not blnInDouble And Not blnInSingle And Not blnInBracket Then
                    lngDepth = lngDepth + 1
                End If
            Case ")"
                If Not blnInDouble And Not blnInSingle And Not blnInBracket Then
                    lngDepth = lngDepth - 1
                End If
        End Select

        If Not blnInDouble And Not blnInSingle And Not blnInBracket And lngDepth = 0 Then
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
    Dim blnInDouble As Boolean
    Dim blnInSingle As Boolean
    Dim blnInBracket As Boolean
    Dim strUpper As String
    Dim strKeyUpper As String

    strUpper = UCase$(strSql)
    strKeyUpper = UCase$(strKeyword)

    For lngPos = lngStart To Len(strUpper) - Len(strKeyUpper) + 1
        Select Case Mid$(strUpper, lngPos, 1)
            Case """"
                If Not blnInSingle And Not blnInBracket Then
                    blnInDouble = Not blnInDouble
                End If
            Case "'"
                If Not blnInDouble And Not blnInBracket Then
                    blnInSingle = Not blnInSingle
                End If
            Case "["
                If Not blnInDouble And Not blnInSingle Then blnInBracket = True
            Case "]"
                If Not blnInDouble And Not blnInSingle Then blnInBracket = False
            Case "("
                If Not blnInDouble And Not blnInSingle And Not blnInBracket Then
                    lngDepth = lngDepth + 1
                End If
            Case ")"
                If Not blnInDouble And Not blnInSingle And Not blnInBracket Then
                    lngDepth = lngDepth - 1
                End If
        End Select

        If Not blnInDouble And Not blnInSingle And Not blnInBracket And lngDepth = 0 Then
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
    Dim blnInDouble As Boolean
    Dim blnInSingle As Boolean
    Dim blnInBracket As Boolean

    lngDepth = 0
    For lngPos = lngStart To Len(strExpr)
        Select Case Mid$(strExpr, lngPos, 1)
            Case """"
                If Not blnInSingle And Not blnInBracket Then
                    blnInDouble = Not blnInDouble
                End If
            Case "'"
                If Not blnInDouble And Not blnInBracket Then
                    blnInSingle = Not blnInSingle
                End If
            Case "["
                If Not blnInDouble And Not blnInSingle Then blnInBracket = True
            Case "]"
                If Not blnInDouble And Not blnInSingle Then blnInBracket = False
            Case "("
                If Not blnInDouble And Not blnInSingle And Not blnInBracket Then
                    lngDepth = lngDepth + 1
                End If
            Case ")"
                If Not blnInDouble And Not blnInSingle And Not blnInBracket Then
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


'---------------------------------------------------------------------------------------
' Procedure : BracketIfNeeded
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Bracket an identifier that is not a plain word or is reserved.
'           : blnDesignerWords selects the word list the Access designer uses
'           : (export format 5.2.0+); otherwise the original shorter list.
'---------------------------------------------------------------------------------------
'
Public Function BracketIfNeeded(strName As String, Optional blnDesignerWords As Boolean = False) As String
    If Left$(strName, 1) = "[" Then
        BracketIfNeeded = strName
    ElseIf InStr(1, strName, ".") > 0 Then
        BracketIfNeeded = strName
    ElseIf HasNonIdentChars(strName) Or IsAccessReservedWord(strName, blnDesignerWords) Then
        BracketIfNeeded = "[" & strName & "]"
    Else
        BracketIfNeeded = strName
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : FormatSetTarget
' Author    : Adam Waller
' Date      : 11/19/2026
' Purpose   : Spell an UPDATE SET target from MSysQueries Name2 so it parses.
'           : Name2 is `Order Lines.Unit Cost` when the query was saved in SQL View
'           : and `[Order Lines].[Unit Cost]` when saved in Design View; both
'           : return the Design View form. Each part keeps its brackets, gains
'           : them only for non-identifier characters, and a qualified reserved
'           : word stays bare, because Access parses it. An unqualified target
'           : also gets brackets for a word on SET_TARGET_RESERVED_WORDS, on top
'           : of BracketIfNeeded's own list, so no bracket an older export wrote
'           : is ever dropped.
'---------------------------------------------------------------------------------------
'
Public Function FormatSetTarget(strTarget As String, Optional blnDesignerWords As Boolean = False) As String
    Dim strQualifier As String
    Dim strMember As String

    If SplitQualifiedName(strTarget, strQualifier, strMember) Then
        FormatSetTarget = BracketNamePart(strQualifier) & "." & BracketNamePart(strMember)
    ElseIf Left$(strTarget, 1) = "[" Then
        FormatSetTarget = strTarget
    ElseIf IsSetTargetReservedWord(strTarget) Then
        FormatSetTarget = "[" & strTarget & "]"
    Else
        FormatSetTarget = BracketIfNeeded(strTarget, blnDesignerWords)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : SplitQualifiedName
' Author    : Adam Waller
' Date      : 11/19/2026
' Purpose   : Split `table.field` at the first period outside brackets. Access does
'           : not allow a period in a table or field name, local or linked, so the
'           : first one always separates the two. Returns False when there is none.
'---------------------------------------------------------------------------------------
'
Public Function SplitQualifiedName(strName As String, ByRef strQualifier As String, _
    ByRef strMember As String) As Boolean

    Dim lngPos As Long
    Dim blnInBracket As Boolean

    For lngPos = 1 To Len(strName)
        Select Case Mid$(strName, lngPos, 1)
            Case "["
                blnInBracket = True
            Case "]"
                blnInBracket = False
            Case "."
                If Not blnInBracket Then
                    strQualifier = Trim$(Left$(strName, lngPos - 1))
                    strMember = Trim$(Mid$(strName, lngPos + 1))
                    SplitQualifiedName = (Len(strQualifier) > 0 And Len(strMember) > 0)
                    Exit Function
                End If
        End Select
    Next lngPos
End Function


'---------------------------------------------------------------------------------------
' Procedure : FindUnquotedChar
' Author    : Adam Waller
' Date      : 11/19/2026
' Purpose   : Position of the first strChar outside quotes and brackets, or 0.
'           : Finds the column list of `INSERT INTO [tblCars (Archive)] (a, b)`
'           : or `INSERT INTO t IN 'C:\path (x)\db.accdb' (a, b)` without
'           : stopping at a parenthesis inside the name or the path.
'---------------------------------------------------------------------------------------
'
Public Function FindUnquotedChar(strSql As String, strChar As String) As Long
    Dim lngPos As Long
    Dim strCh As String
    Dim blnInDouble As Boolean
    Dim blnInSingle As Boolean
    Dim blnInBracket As Boolean

    For lngPos = 1 To Len(strSql)
        strCh = Mid$(strSql, lngPos, 1)
        If blnInDouble Then
            If strCh = """" Then blnInDouble = False
        ElseIf blnInSingle Then
            If strCh = "'" Then blnInSingle = False
        ElseIf blnInBracket Then
            If strCh = "]" Then blnInBracket = False
        ElseIf strCh = strChar Then
            FindUnquotedChar = lngPos
            Exit Function
        Else
            Select Case strCh
                Case """": blnInDouble = True
                Case "'": blnInSingle = True
                Case "[": blnInBracket = True
            End Select
        End If
    Next lngPos
End Function


Public Function IsSetTargetReservedWord(strWord As String) As Boolean
    If Len(strWord) = 0 Or HasNonIdentChars(strWord) Then Exit Function
    IsSetTargetReservedWord = (InStr(1, SET_TARGET_RESERVED_WORDS, " " & UCase$(strWord) & " ", vbBinaryCompare) > 0)
End Function


Private Function BracketNamePart(strPart As String) As String
    If Left$(strPart, 1) = "[" Or Not HasNonIdentChars(strPart) Then
        BracketNamePart = strPart
    Else
        BracketNamePart = "[" & strPart & "]"
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


Public Function IsAccessReservedWord(strWord As String, Optional blnDesignerWords As Boolean = False) As Boolean
    If blnDesignerWords Then
        If Len(strWord) = 0 Or HasNonIdentChars(strWord) Then Exit Function
        IsAccessReservedWord = (InStr(1, DESIGNER_RESERVED_WORDS, " " & UCase$(strWord) & " ", vbBinaryCompare) > 0)
        Exit Function
    End If
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
