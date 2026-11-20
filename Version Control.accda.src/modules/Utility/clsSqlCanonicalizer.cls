VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSqlCanonicalizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsSqlCanonicalizer
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Remove provably redundant parentheses from the WHERE and HAVING clauses
'           : of Access SQL, so criteria saved by the query designer and criteria
'           : written in SQL View export the same way. Only parenthesis tokens are
'           : removed; every other character is kept as written, and layout is left
'           : to clsSqlFormatter. This class does not read Options: the caller
'           : decides whether to run it.
'           :
'           : Rules, applied until nothing changes:
'           :   1. A pair wrapping the entire WHERE or HAVING condition.
'           :   2. A pair around a single value: a bare, bracketed, or qualified name
'           :      (including Forms!frm!ctl), a bracketed parameter, a string, a
'           :      #date#, or an unsigned number. Repeated wrappers collapse the same way.
'           : No operator precedence is used. Function calls, IN, EXISTS, ANY/ALL/SOME,
'           : subqueries, and everything outside WHERE/HAVING (FROM, JOIN/ON) keep their
'           : parentheses. If the result would change any non-parenthesis token or
'           : leave parentheses unbalanced, the input is returned unchanged and
'           : DeclineReason says why.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Utility")

Private m_lngCount As Long
Private m_intType() As eTokenTypes
Private m_strValue() As String
Private m_blnRemoved() As Boolean
Private m_lngMatch() As Long
Private m_lngRegion() As Long

' WHERE/HAVING regions: keyword token, first and last token of the condition
Private m_lngRegionCount As Long
Private m_lngRegionKeyword() As Long
Private m_lngRegionStart() As Long
Private m_lngRegionEnd() As Long

Private m_blnChanged As Boolean
Private m_strDeclineReason As String


'---------------------------------------------------------------------------------------
' Procedure : Canonicalize
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Return strSql with redundant WHERE/HAVING parentheses removed, or strSql
'           : unchanged when nothing applies or the self-check declines the rewrite.
'---------------------------------------------------------------------------------------
'
Public Function Canonicalize(strSql As String) As String

    Dim lngRegion As Long
    Dim blnRemoved As Boolean
    Dim strResult As String
    Dim strReason As String

    m_blnChanged = False
    m_strDeclineReason = vbNullString
    Canonicalize = strSql
    If Len(strSql) = 0 Then Exit Function

    ' Own the error state rather than inheriting the caller's On Error Resume Next,
    ' which could return a partly rewritten statement.
    LogUnhandledErrors
    On Error GoTo ErrHandler

    LoadTokens strSql
    If HasUnsupportedTokens Then Exit Function
    If Not MatchParentheses Then
        m_strDeclineReason = "unbalanced parentheses in input"
        Exit Function
    End If

    FindRegions
    If m_lngRegionCount = 0 Then Exit Function

    Do
        blnRemoved = False
        For lngRegion = 1 To m_lngRegionCount
            If RemoveClauseWrapper(lngRegion) Then blnRemoved = True
            If RemoveSingleValueWrappers(lngRegion) Then blnRemoved = True
        Next lngRegion
    Loop While blnRemoved

    strResult = BuildOutput
    If strResult = strSql Then Exit Function

    strReason = CheckRewrite(strSql, strResult)
    If Len(strReason) Then
        m_strDeclineReason = strReason
        Exit Function
    End If

    m_blnChanged = True
    Canonicalize = strResult
    Exit Function

ErrHandler:
    m_strDeclineReason = "error " & Err.Number & " " & Err.Description
    Err.Clear
    m_blnChanged = False
    Canonicalize = strSql

End Function


' True when the last call to Canonicalize removed any parentheses.
Public Property Get Changed() As Boolean
    Changed = m_blnChanged
End Property


' Why the last call to Canonicalize kept its input; empty when it did not decline.
Public Property Get DeclineReason() As String
    DeclineReason = m_strDeclineReason
End Property


'---------------------------------------------------------------------------------------
' Procedure : CheckRewrite
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : The fail-safe. Return an empty string when strAfter differs from
'           : strBefore only by removed parentheses and whitespace, and its
'           : parentheses are balanced; otherwise return the reason it does not.
'---------------------------------------------------------------------------------------
'
Friend Function CheckRewrite(strBefore As String, strAfter As String) As String

    Dim colBefore As Collection
    Dim colAfter As Collection
    Dim lngItem As Long
    Dim lngDepth As Long
    Dim varToken As Variant

    Set colBefore = SignificantTokens(strBefore)
    Set colAfter = SignificantTokens(strAfter)

    If colBefore.Count <> colAfter.Count Then
        CheckRewrite = "non-parenthesis tokens changed"
        Exit Function
    End If
    For lngItem = 1 To colBefore.Count
        If StrComp(colBefore(lngItem), colAfter(lngItem), vbBinaryCompare) <> 0 Then
            CheckRewrite = "non-parenthesis tokens changed"
            Exit Function
        End If
    Next lngItem

    With New clsSqlFormatter
        .Tokenize strAfter, esdAccess
        For Each varToken In .Tokens
            If varToken(0) = ttBoundary Then
                If varToken(1) = "(" Then
                    lngDepth = lngDepth + 1
                ElseIf varToken(1) = ")" Then
                    lngDepth = lngDepth - 1
                    If lngDepth < 0 Then Exit For
                End If
            End If
        Next varToken
    End With
    If lngDepth <> 0 Then CheckRewrite = "unbalanced parentheses in result"

End Function


'---------------------------------------------------------------------------------------
' Procedure : SignificantTokens
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Return "type|value" for every token that is not whitespace or a
'           : parenthesis, in order.
'---------------------------------------------------------------------------------------
'
Private Function SignificantTokens(strSql As String) As Collection

    Dim varToken As Variant

    Set SignificantTokens = New Collection
    With New clsSqlFormatter
        .Tokenize strSql, esdAccess
        For Each varToken In .Tokens
            If varToken(0) <> ttWhitespace Then
                If Not IsParenToken(CLng(varToken(0)), CStr(varToken(1))) Then
                    SignificantTokens.Add varToken(0) & "|" & varToken(1)
                End If
            End If
        Next varToken
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadTokens
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Tokenize strSql with the formatter's tokenizer into parallel arrays.
'---------------------------------------------------------------------------------------
'
Private Sub LoadTokens(strSql As String)

    Dim varToken As Variant
    Dim lngToken As Long

    With New clsSqlFormatter
        .Tokenize strSql, esdAccess
        m_lngCount = .Tokens.Count
        ReDim m_intType(0 To m_lngCount + 1)
        ReDim m_strValue(0 To m_lngCount + 1)
        ReDim m_blnRemoved(0 To m_lngCount + 1)
        ReDim m_lngMatch(0 To m_lngCount + 1)
        ReDim m_lngRegion(0 To m_lngCount + 1)
        For Each varToken In .Tokens
            lngToken = lngToken + 1
            m_intType(lngToken) = varToken(0)
            m_strValue(lngToken) = varToken(1)
        Next varToken
    End With

    m_lngRegionCount = 0

End Sub


'---------------------------------------------------------------------------------------
' Procedure : HasUnsupportedTokens
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Saved Access SQL never contains comments or text the tokenizer cannot
'           : read; anything that does is left as written rather than risk
'           : misreading where a clause ends.
'---------------------------------------------------------------------------------------
'
Private Function HasUnsupportedTokens() As Boolean
    Dim lngToken As Long
    For lngToken = 1 To m_lngCount
        Select Case m_intType(lngToken)
            Case ttComment, ttBlockComment, ttUnknown, ttError
                HasUnsupportedTokens = True
                Exit Function
        End Select
    Next lngToken
End Function


'---------------------------------------------------------------------------------------
' Procedure : MatchParentheses
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Record the matching token for each parenthesis. Returns False when the
'           : parentheses are unbalanced.
'---------------------------------------------------------------------------------------
'
Private Function MatchParentheses() As Boolean

    Dim lngStack() As Long
    Dim lngDepth As Long
    Dim lngToken As Long

    ReDim lngStack(1 To m_lngCount + 1)
    For lngToken = 1 To m_lngCount
        If IsParenToken(m_intType(lngToken), m_strValue(lngToken)) Then
            If m_strValue(lngToken) = "(" Then
                lngDepth = lngDepth + 1
                lngStack(lngDepth) = lngToken
            Else
                If lngDepth = 0 Then Exit Function
                m_lngMatch(lngToken) = lngStack(lngDepth)
                m_lngMatch(lngStack(lngDepth)) = lngToken
                lngDepth = lngDepth - 1
            End If
        End If
    Next lngToken

    MatchParentheses = (lngDepth = 0)

End Function


'---------------------------------------------------------------------------------------
' Procedure : FindRegions
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Locate each WHERE and HAVING condition, including those in subqueries
'           : and UNION branches, and tag every token with the innermost region that
'           : owns it. Tokens inside a nested subquery belong to that subquery's own
'           : region (or to none), never to the enclosing condition.
'---------------------------------------------------------------------------------------
'
Private Sub FindRegions()

    Dim lngToken As Long
    Dim lngEnd As Long
    Dim lngRegion As Long

    ReDim m_lngRegionKeyword(1 To m_lngCount)
    ReDim m_lngRegionStart(1 To m_lngCount)
    ReDim m_lngRegionEnd(1 To m_lngCount)

    For lngToken = 1 To m_lngCount
        If m_intType(lngToken) = ttReservedTopLevel Then
            Select Case UCase$(m_strValue(lngToken))
                Case "WHERE", "HAVING"
                    lngEnd = FindClauseEnd(lngToken)
                    If lngEnd > lngToken Then
                        m_lngRegionCount = m_lngRegionCount + 1
                        m_lngRegionKeyword(m_lngRegionCount) = lngToken
                        m_lngRegionStart(m_lngRegionCount) = lngToken + 1
                        m_lngRegionEnd(m_lngRegionCount) = lngEnd
                    End If
            End Select
        End If
    Next lngToken

    ' Outer regions come first, so inner regions claim their own tokens afterwards.
    For lngRegion = 1 To m_lngRegionCount
        lngToken = m_lngRegionStart(lngRegion)
        Do While lngToken <= m_lngRegionEnd(lngRegion)
            m_lngRegion(lngToken) = lngRegion
            If m_strValue(lngToken) = "(" And IsParenToken(m_intType(lngToken), m_strValue(lngToken)) Then
                If IsSubqueryGroup(lngToken) Then
                    m_lngRegion(m_lngMatch(lngToken)) = lngRegion
                    lngToken = m_lngMatch(lngToken)
                End If
            End If
            lngToken = lngToken + 1
        Loop
    Next lngRegion

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FindClauseEnd
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Return the last token of the condition after the WHERE/HAVING keyword
'           : at lngKeyword, or 0 when the clause cannot be bounded with confidence.
'---------------------------------------------------------------------------------------
'
Private Function FindClauseEnd(lngKeyword As Long) As Long

    Dim lngToken As Long
    Dim lngLast As Long

    lngToken = lngKeyword
    Do
        lngToken = lngToken + 1
        If lngToken > m_lngCount Then Exit Do
        If m_intType(lngToken) = ttWhitespace Then
            ' Skip
        ElseIf m_intType(lngToken) = ttReservedTopLevel Then
            Select Case UCase$(m_strValue(lngToken))
                Case "GROUP BY", "ORDER BY", "HAVING", "UNION", "UNION ALL", "PIVOT", "WITH"
                    Exit Do
                Case Else
                    ' A clause keyword where none belongs (perhaps an unbracketed
                    ' field name). Leave this condition as written.
                    Exit Function
            End Select
        ElseIf IsParenToken(m_intType(lngToken), m_strValue(lngToken)) Then
            If m_strValue(lngToken) = ")" Then Exit Do
            lngToken = m_lngMatch(lngToken)
            lngLast = lngToken
        ElseIf m_strValue(lngToken) = ";" Then
            Exit Do
        Else
            lngLast = lngToken
        End If
    Loop

    FindClauseEnd = lngLast

End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveClauseWrapper
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Rule 1. Remove a pair that wraps the entire condition of a region.
'---------------------------------------------------------------------------------------
'
Private Function RemoveClauseWrapper(lngRegion As Long) As Boolean

    Dim lngFirst As Long
    Dim lngLast As Long

    lngFirst = NextSignificant(m_lngRegionStart(lngRegion) - 1)
    lngLast = PreviousSignificant(m_lngRegionEnd(lngRegion) + 1)
    If lngFirst = 0 Or lngLast = 0 Then Exit Function
    If lngFirst > m_lngRegionEnd(lngRegion) Then Exit Function
    If m_strValue(lngFirst) <> "(" Then Exit Function
    If Not IsParenToken(m_intType(lngFirst), m_strValue(lngFirst)) Then Exit Function
    If m_lngMatch(lngFirst) <> lngLast Then Exit Function
    If IsSubqueryGroup(lngFirst) Then Exit Function

    m_blnRemoved(lngFirst) = True
    m_blnRemoved(lngLast) = True
    RemoveClauseWrapper = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveSingleValueWrappers
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Rule 2. Remove each pair in the region that wraps a single value and
'           : sits where a plain operand may stand.
'---------------------------------------------------------------------------------------
'
Private Function RemoveSingleValueWrappers(lngRegion As Long) As Boolean

    Dim lngToken As Long
    Dim lngClose As Long

    For lngToken = m_lngRegionStart(lngRegion) To m_lngRegionEnd(lngRegion)
        If Not m_blnRemoved(lngToken) And m_lngRegion(lngToken) = lngRegion Then
            If m_strValue(lngToken) = "(" And IsParenToken(m_intType(lngToken), m_strValue(lngToken)) Then
                lngClose = m_lngMatch(lngToken)
                If IsSingleValue(lngToken, lngClose) Then
                    If OperandAllowedAfter(lngToken, lngRegion) And OperandAllowedBefore(lngClose) Then
                        m_blnRemoved(lngToken) = True
                        m_blnRemoved(lngClose) = True
                        RemoveSingleValueWrappers = True
                    End If
                End If
            End If
        End If
    Next lngToken

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsSingleValue
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Return True when the tokens between lngOpen and lngClose form exactly one
'           : atomic operand. Anything not positively recognized returns False.
'---------------------------------------------------------------------------------------
'
Private Function IsSingleValue(lngOpen As Long, lngClose As Long) As Boolean

    Dim lngToken As Long
    Dim lngParts As Long
    Dim blnExpectName As Boolean

    lngToken = NextSignificant(lngOpen)
    If lngToken = 0 Or lngToken >= lngClose Then Exit Function

    ' A lone string, date, bracketed name or parameter, or unsigned number
    If NextSignificant(lngToken) = lngClose Then
        Select Case m_intType(lngToken)
            Case ttQuote, ttNumber, ttWord
                IsSingleValue = True
        End Select
        Exit Function
    End If

    ' A qualified name: name ( ("." | "!") name )*
    blnExpectName = True
    Do While lngToken < lngClose
        If blnExpectName Then
            Select Case m_intType(lngToken)
                Case ttWord
                Case ttQuote
                    ' Only bracketed names qualify, not strings or dates
                    If Left$(m_strValue(lngToken), 1) <> "[" Then Exit Function
                Case Else
                    Exit Function
            End Select
            lngParts = lngParts + 1
        Else
            If m_strValue(lngToken) <> "." And m_strValue(lngToken) <> "!" Then Exit Function
        End If
        blnExpectName = Not blnExpectName
        lngToken = NextSignificant(lngToken)
        If lngToken = 0 Then Exit Function
    Loop

    IsSingleValue = (Not blnExpectName) And (lngParts > 1)

End Function


'---------------------------------------------------------------------------------------
' Procedure : OperandAllowedAfter
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Return True when the token before the "(" at lngOpen lets a plain operand
'           : stand in its place: the WHERE/HAVING keyword, an operator, a logical
'           : keyword, or another grouping parenthesis that is itself allowed. A "("
'           : after a name is a function call; after IN, EXISTS, ANY, ALL, SOME, or
'           : VALUES it is part of that syntax.
'---------------------------------------------------------------------------------------
'
Private Function OperandAllowedAfter(lngOpen As Long, lngRegion As Long) As Boolean

    Dim lngPrev As Long

    lngPrev = PreviousSignificant(lngOpen)
    If lngPrev = 0 Then Exit Function

    If lngPrev = m_lngRegionKeyword(lngRegion) Then
        OperandAllowedAfter = True
        Exit Function
    End If

    Select Case m_intType(lngPrev)
        Case ttBoundary
            Select Case m_strValue(lngPrev)
                Case "=", "<", ">", "+", "-", "*", "/", "&", "^"
                    OperandAllowedAfter = True
                Case "("
                    If m_lngRegion(lngPrev) = lngRegion And Not IsSubqueryGroup(lngPrev) Then
                        OperandAllowedAfter = OperandAllowedAfter(lngPrev, lngRegion)
                    End If
            End Select
        Case ttReserved, ttReservedNewline, ttWord
            Select Case UCase$(m_strValue(lngPrev))
                Case "AND", "OR", "XOR", "NOT", "EQV", "IMP", "MOD", "LIKE", "BETWEEN", "\"
                    OperandAllowedAfter = True
            End Select
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : OperandAllowedBefore
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Return False when the token after the ")" at lngClose continues the
'           : group as a name or call, as in `(x).y`, `(x)!y`, or `(x)(y)`.
'---------------------------------------------------------------------------------------
'
Private Function OperandAllowedBefore(lngClose As Long) As Boolean

    Dim lngNext As Long

    lngNext = NextSignificant(lngClose)
    If lngNext = 0 Then
        OperandAllowedBefore = True
    Else
        Select Case m_strValue(lngNext)
            Case ".", "!", "("
            Case Else
                OperandAllowedBefore = True
        End Select
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsSubqueryGroup
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Return True when the group opened at lngOpen starts with a clause
'           : keyword such as SELECT.
'---------------------------------------------------------------------------------------
'
Private Function IsSubqueryGroup(lngOpen As Long) As Boolean
    Dim lngNext As Long
    lngNext = NextSignificant(lngOpen)
    If lngNext > 0 Then IsSubqueryGroup = (m_intType(lngNext) = ttReservedTopLevel)
End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildOutput
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Join the remaining tokens. A removed parenthesis becomes a space only
'           : where its neighbors would otherwise run together into one token.
'---------------------------------------------------------------------------------------
'
Private Function BuildOutput() As String

    Dim lngToken As Long
    Dim lngNext As Long
    Dim strLast As String
    Dim strNext As String

    With New clsConcat
        For lngToken = 1 To m_lngCount
            If m_blnRemoved(lngToken) Then
                lngNext = lngToken + 1
                Do While lngNext <= m_lngCount
                    If Not m_blnRemoved(lngNext) Then Exit Do
                    lngNext = lngNext + 1
                Loop
                If lngNext <= m_lngCount Then strNext = Left$(m_strValue(lngNext), 1) Else strNext = vbNullString
                If Not (IsSeparatorChar(strLast) Or IsSeparatorChar(strNext)) Then
                    .Add " "
                    strLast = " "
                End If
            Else
                .Add m_strValue(lngToken)
                strLast = Right$(m_strValue(lngToken), 1)
            End If
        Next lngToken
        BuildOutput = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsSeparatorChar
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Return True for an empty string, whitespace, or punctuation that ends a
'           : token on its own.
'---------------------------------------------------------------------------------------
'
Private Function IsSeparatorChar(strChar As String) As Boolean
    If Len(strChar) = 0 Then
        IsSeparatorChar = True
    Else
        IsSeparatorChar = (InStr(1, " " & vbTab & vbCr & vbLf & ",;:)(.=<>+-*/!^%|&", strChar, vbBinaryCompare) > 0)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : NextSignificant
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Return the next token after lngToken that is not whitespace and has not
'           : been removed, or 0 at the end.
'---------------------------------------------------------------------------------------
'
Private Function NextSignificant(lngToken As Long) As Long
    Dim lngNext As Long
    For lngNext = lngToken + 1 To m_lngCount
        If m_intType(lngNext) <> ttWhitespace And Not m_blnRemoved(lngNext) Then
            NextSignificant = lngNext
            Exit Function
        End If
    Next lngNext
End Function


'---------------------------------------------------------------------------------------
' Procedure : PreviousSignificant
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Return the previous token before lngToken that is not whitespace and has
'           : not been removed, or 0 at the start.
'---------------------------------------------------------------------------------------
'
Private Function PreviousSignificant(lngToken As Long) As Long
    Dim lngPrev As Long
    For lngPrev = lngToken - 1 To 1 Step -1
        If m_intType(lngPrev) <> ttWhitespace And Not m_blnRemoved(lngPrev) Then
            PreviousSignificant = lngPrev
            Exit Function
        End If
    Next lngPrev
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsParenToken
' Author    : Adam Waller
' Date      : 10/28/2026
' Purpose   : Return True for a "(" or ")" punctuation token.
'---------------------------------------------------------------------------------------
'
Private Function IsParenToken(ByVal intType As eTokenTypes, ByVal strValue As String) As Boolean
    If intType = ttBoundary Then IsParenToken = (strValue = "(" Or strValue = ")")
End Function
