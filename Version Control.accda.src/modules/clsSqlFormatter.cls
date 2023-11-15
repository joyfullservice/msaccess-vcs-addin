VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSqlFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsSqlFormatter
' Author    : Adam Waller
' Date      : 8/11/2023
' Purpose   : Provide a "simple" implementation of SQL formatting for consistent output
'           : in source code files. This is largely based on the doctrine/sql-formatter
'           : project on GitHub. For my purposes I didn't need a huge multi-dialect
'           : parser with lots of output options, but something that I could inlcude
'           : as a single module in a project to parse SQL statements from Access
'           : queries, as well as basic support for other common dialects like MSSQL
'           : and MySQL.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


' Lists Updated 8/11/2023 from doctrine/sql-formatter
Private Const cstrReserved As String = _
    "|ACCESSIBLE|ACTION|AFTER|AGAINST|AGGREGATE|ALGORITHM|ALL|ALTER|ANALYSE|ANALYZE|AS|ASC|AUTOCOMMIT|AUTO_INCREMENT|BACKUP|BEGIN|BETWEEN|BINLOG|BOTH|CASCADE|CASE|CHANGE|CHANGED|CHARACTER SET|CHARSET|CHECK|CHECKSUM|COLLATE|COLLATION|COLUMN|COLUMNS|COMMENT|COMMIT|COMMITTED|COMPRESSED|CONCURRENT|CONSTRAINT|CONTAINS|CONVERT|CREATE|CROSS|CURRENT ROW|CURRENT_TIMESTAMP|DATABASE|DATABASES|DAY|DAY_HOUR|DAY_MINUTE|DAY_SECOND|DEFAULT|DEFINER|DELAYED|DELETE|DESC|DESCRIBE|DETERMINISTIC|DISTINCT|DISTINCTROW|DIV|DO|DUMPFILE|DUPLICATE|DYNAMIC|ELSE|ENCLOSED|END|ENGINE|ENGINE_TYPE|ENGINES|ESCAPE|ESCAPED|EVENTS|EXEC|EXECUTE|EXISTS|EXPLAIN|EXTENDED|FAST|FIELDS|FILE|FILTER|FIRST|FIXED|FLUSH|FOR|FORCE|FOLLOWING|FOREIGN|FULL|FULLTEXT|FUNCTION|GLOBAL|GRANT|GRANTS|GROUP|GROUPS|HEAP|HIGH_PRIORITY|HOSTS|HOUR|HOUR_MINUTE|HOUR_SECOND|IDENTIFIED|IF|IFNULL|IGNORE|IN|INDEX|INDEXES|INFILE|INSERT|INSERT_ID|INSERT_METHOD|INTERVAL|INTO|INVOKER|IS|ISOLATION|" & _
    "KEY|KEYS|KILL|LAST_INSERT_ID|LEADING|LEVEL|LIKE|LINEAR|LINES|LOAD|LOCAL|LOCK|LOCKS|LOGS|LOW_PRIORITY|MARIA|MASTER|MASTER_CONNECT_RETRY|MASTER_HOST|MASTER_LOG_FILE|MATCH|MAX_CONNECTIONS_PER_HOUR|MAX_QUERIES_PER_HOUR|MAX_ROWS|MAX_UPDATES_PER_HOUR|MAX_USER_CONNECTIONS|MEDIUM|MERGE|MINUTE|MINUTE_SECOND|MIN_ROWS|MODE|MONTH|MRG_MYISAM|MYISAM|NAMES|NATURAL|NO OTHERS|NOT|NOW()|NULL|OFFSET|ON|OPEN|OPTIMIZE|OPTION|OPTIONALLY|ON UPDATE|ON DELETE|OUTFILE|OVER|PACK_KEYS|PAGE|PARTIAL|PARTITION|PARTITIONS|PASSWORD|PRECEDING|PRIMARY|PRIVILEGES|PROCEDURE|PROCESS|PROCESSLIST|PURGE|QUICK|RANGE|RAID0|RAID_CHUNKS|RAID_CHUNKSIZE|RAID_TYPE|READ|READ_ONLY|READ_WRITE|RECURSIVE|REFERENCES|REGEXP|RELOAD|RENAME|REPAIR|REPEATABLE|REPLACE|REPLICATION|RESET|RESTORE|RESTRICT|RETURN|RETURNS|REVOKE|RLIKE|ROLLBACK|ROW|ROWS|ROW_FORMAT|SECOND|SECURITY|SEPARATOR|SERIALIZABLE|SESSION|SHARE|SHOW|SHUTDOWN|SLAVE|SONAME|SOUNDS|SQL|SQL_AUTO_IS_NULL|SQL_BIG_RESULT|" & _
    "SQL_BIG_SELECTS|SQL_BIG_TABLES|SQL_BUFFER_RESULT|SQL_CALC_FOUND_ROWS|SQL_LOG_BIN|SQL_LOG_OFF|SQL_LOG_UPDATE|SQL_LOW_PRIORITY_UPDATES|SQL_MAX_JOIN_SIZE|SQL_QUOTE_SHOW_CREATE|SQL_SAFE_UPDATES|SQL_SELECT_LIMIT|SQL_SLAVE_SKIP_COUNTER|SQL_SMALL_RESULT|SQL_WARNINGS|SQL_CACHE|SQL_NO_CACHE|START|STARTING|STATUS|STOP|STORAGE|STRAIGHT_JOIN|STRING|STRIPED|SUPER|TABLE|TABLES|TEMPORARY|TERMINATED|THEN|TIES|TO|TRAILING|TRANSACTIONAL|TRUE|TRUNCATE|TYPE|TYPES|UNBOUNDED|UNCOMMITTED|UNIQUE|UNLOCK|UNSIGNED|USAGE|USE|USING|VARIABLES|VIEW|WHEN|WITH|WORK|WRITE|YEAR_MONTH|"
Private Const cstrFunctions As String = _
    "|ABS|ACOS|ADDDATE|ADDTIME|AES_DECRYPT|AES_ENCRYPT|APPROX_COUNT_DISTINCT|AREA|ASBINARY|ASCII|ASIN|ASTEXT|ATAN|ATAN2|AVG|BDMPOLYFROMTEXT|BDMPOLYFROMWKB|BDPOLYFROMTEXT|BDPOLYFROMWKB|BENCHMARK|BIN|BIT_AND|BIT_COUNT|BIT_LENGTH|BIT_OR|BIT_XOR|BOUNDARY|BUFFER|CAST|CEIL|CEILING|CENTROID|CHAR|CHARACTER_LENGTH|CHARSET|CHAR_LENGTH|CHECKSUM_AGG|COALESCE|COERCIBILITY|COLLATION|COMPRESS|CONCAT|CONCAT_WS|CONNECTION_ID|CONTAINS|CONV|CONVERT|CONVERT_TZ|CONVEXHULL|COS|COT|COUNT|COUNT_BIG|CRC32|CROSSES|CUME_DIST|CURDATE|CURRENT_DATE|CURRENT_TIME|CURRENT_TIMESTAMP|CURRENT_USER|CURTIME|DATABASE|DATE|DATEDIFF|DATE_ADD|DATE_DIFF|DATE_FORMAT|DATE_SUB|DAY|DAYNAME|DAYOFMONTH|DAYOFWEEK|DAYOFYEAR|DECODE|DEFAULT|DEGREES|DENSE_RANK|DES_DECRYPT|DES_ENCRYPT|DIFFERENCE|DIMENSION|DISJOINT|DISTANCE|ELT|ENCODE|ENCRYPT|ENDPOINT|ENVELOPE|EQUALS|EXP|EXPORT_SET|EXTERIORRING|EXTRACT|EXTRACTVALUE|FIELD|FIND_IN_SET|FIRST_VALUE|FLOOR|FORMAT|FOUND_ROWS|FROM_DAYS|" & _
    "FROM_UNIXTIME|GEOMCOLLFROMTEXT|GEOMCOLLFROMWKB|GEOMETRYCOLLECTION|GEOMETRYCOLLECTIONFROMTEXT|GEOMETRYCOLLECTIONFROMWKB|GEOMETRYFROMTEXT|GEOMETRYFROMWKB|GEOMETRYN|GEOMETRYTYPE|GEOMFROMTEXT|GEOMFROMWKB|GET_FORMAT|GET_LOCK|GLENGTH|GREATEST|GROUPING|GROUPING_ID|GROUP_CONCAT|GROUP_UNIQUE_USERS|HEX|HOUR|IF|IFNULL|INET_ATON|INET_NTOA|INSERT|INSTR|INTERIORRINGN|INTERSECTION|INTERSECTS|INTERVAL|ISCLOSED|ISEMPTY|ISNULL|ISRING|ISSIMPLE|IS_FREE_LOCK|IS_USED_LOCK|LAG|LAST_DAY|LAST_INSERT_ID|LAST_VALUE|LCASE|LEAD|LEAST|LEFT|LENGTH|LINEFROMTEXT|LINEFROMWKB|LINESTRING|LINESTRINGFROMTEXT|LINESTRINGFROMWKB|LISTAGG|LN|LOAD_FILE|LOCALTIME|LOCALTIMESTAMP|LOCATE|LOG|LOG10|LOG2|LOWER|LPAD|LTRIM|MAKEDATE|MAKETIME|MAKE_SET|MASTER_POS_WAIT|MAX|MBRCONTAINS|MBRDISJOINT|MBREQUAL|MBRINTERSECTS|MBROVERLAPS|MBRTOUCHES|MBRWITHIN|MD5|MICROSECOND|MID|MIN|MINUTE|MLINEFROMTEXT|MLINEFROMWKB|MOD|MONTH|MONTHNAME|MPOINTFROMTEXT|MPOINTFROMWKB|MPOLYFROMTEXT|" & _
    "MPOLYFROMWKB|MULTILINESTRING|MULTILINESTRINGFROMTEXT|MULTILINESTRINGFROMWKB|MULTIPOINT|MULTIPOINTFROMTEXT|MULTIPOINTFROMWKB|MULTIPOLYGON|MULTIPOLYGONFROMTEXT|MULTIPOLYGONFROMWKB|NAME_CONST|NTH_VALUE|NTILE|NULLIF|NUMGEOMETRIES|NUMINTERIORRINGS|NUMPOINTS|OCT|OCTET_LENGTH|OLD_PASSWORD|ORD|OVERLAPS|PASSWORD|PERCENT_RANK|PERCENTILE_CONT|PERCENTILE_DISC|PERIOD_ADD|PERIOD_DIFF|PI|POINT|POINTFROMTEXT|POINTFROMWKB|POINTN|POINTONSURFACE|POLYFROMTEXT|POLYFROMWKB|POLYGON|POLYGONFROMTEXT|POLYGONFROMWKB|POSITION|POW|POWER|QUARTER|QUOTE|RADIANS|RAND|RANK|RELATED|RELEASE_LOCK|REPEAT|REPLACE|REVERSE|RIGHT|ROUND|ROW_COUNT|ROW_NUMBER|RPAD|RTRIM|SCHEMA|SECOND|SEC_TO_TIME|SESSION_USER|SHA|SHA1|SIGN|SIN|SLEEP|SOUNDEX|SPACE|SQRT|SRID|STARTPOINT|STD|STDEV|STDEVP|STDDEV|STDDEV_POP|STDDEV_SAMP|STRING_AGG|STRCMP|STR_TO_DATE|SUBDATE|SUBSTR|SUBSTRING|SUBSTRING_INDEX|SUBTIME|SUM|SYMDIFFERENCE|SYSDATE|SYSTEM_USER|TAN|TIME|TIMEDIFF|TIMESTAMP|TIMESTAMPADD|" & _
    "TIMESTAMPDIFF|TIME_FORMAT|TIME_TO_SEC|TOUCHES|TO_DAYS|TRIM|TRUNCATE|UCASE|UNCOMPRESS|UNCOMPRESSED_LENGTH|UNHEX|UNIQUE_USERS|UNIX_TIMESTAMP|UPDATEXML|UPPER|USER|UTC_DATE|UTC_TIME|UTC_TIMESTAMP|UUID|VAR|VARIANCE|VARP|VAR_POP|VAR_SAMP|VERSION|WEEK|WEEKDAY|WEEKOFYEAR|WITHIN|X|Y|YEAR|YEARWEEK|"
Private Const cstrReservedToplevel As String = "|WITH|SELECT|FROM|WHERE|SET|ORDER BY|GROUP BY|LIMIT|DROP|VALUES|UPDATE|HAVING|ADD|CHANGE|MODIFY|ALTER TABLE|DELETE FROM|UNION ALL|UNION|EXCEPT|INTERSECT|PARTITION BY|ROWS|RANGE|GROUPS|WINDOW|"
Private Const cstrReservedNewline As String = "|LEFT OUTER JOIN|RIGHT OUTER JOIN|LEFT JOIN|RIGHT JOIN|OUTER JOIN|INNER JOIN|JOIN|XOR|OR|AND|EXCLUDE|"
Private Const cstrBoundaries As String = ",;:)(.=<>+-*/!^%|&#"
Private Const cstrRegExSpecial As String = ".\+*?[^]$(){}=!<>|:-#/"

' SQL Dialects
Public Enum eSqlDialect
    esdUnknown
    esdAccess
    esdMSSQL
    esdMySQL
End Enum

' Types of lists
Private Enum eListType
    eltReserved
    eltFunctions
    eltReservedToplevel
    eltReservedNewline
    eltBoundaries
    eltRegExSpecial
End Enum

' Token types
Private Enum eTokenTypes
    ttUnknown
    ttWhitespace
    ttWord
    ttQuote
    ttBacktickQuote
    ttReserved
    ttReservedTopLevel
    ttReservedNewline
    ttBoundary
    ttComment
    ttBlockComment
    ttNumber
    ttError
    ttVariable
End Enum

Private m_strSql As String
Private m_colTokens As Collection
Private m_lngPos As Long
Private m_intDialect As eSqlDialect
Private m_varWordCache(1 To 2) As Variant


'---------------------------------------------------------------------------------------
' Procedure : FormatSQL
' Author    : Adam Waller
' Date      : 4/1/2020
' Purpose   : This is the main function used outside the class for SQL formatting.
'---------------------------------------------------------------------------------------
'
Public Function FormatSQL(Optional strSql As String, Optional intDialect As eSqlDialect) As String

    Dim lngIndentLevel As Long
    Dim blnNewline As Boolean
    Dim blnInlineParentheses As Boolean
    Dim blnIncreaseSpecialIndent As Boolean
    Dim blnIncreateBlockIndent As Boolean
    Dim colIndents As Collection
    Dim strIndent As String
    Dim blnAddedNewline As Boolean
    Dim intInlineCount As Integer
    Dim blnInlineIndented As Boolean
    Dim blnClauseLimit As Boolean
    Dim lngToken As Long
    Dim lngToken2 As Long
    Dim strNextValue As String
    Dim intNextType As eTokenTypes
    Dim cReturn As clsConcat
    Dim intTokenType As eTokenTypes
    Dim strTokenValue As String
    Dim lngLength As Long
    Dim varItem As Variant
    Dim strType As String

    Perf.CategoryStart "Format SQL"
    Perf.OperationStart "Formating"

    ' Tokenize the string, if provided
    If strSql <> vbNullString Then Tokenize strSql, intDialect

    ' Set up collection to hold types of indents
    Set colIndents = New Collection
    Set cReturn = New clsConcat

    ' Build formatted output from tokens
    For lngToken = 1 To m_colTokens.Count

        intTokenType = m_colTokens(lngToken)(0)
        strTokenValue = m_colTokens(lngToken)(1)

        ' Only process non-whitespace tokens
        If intTokenType = ttWhitespace Then GoTo NextToken

        ' If we are increasing the special indent level
        If blnIncreaseSpecialIndent Then
            IncreaseIndent colIndents, lngIndentLevel, "special"
            blnIncreaseSpecialIndent = False
        End If

        ' If we are increasing the block indent level
        If blnIncreateBlockIndent Then
            IncreaseIndent colIndents, lngIndentLevel, "block"
            blnIncreateBlockIndent = False
        End If

        ' If we need a new line before the token
        If blnNewline Then
            cReturn.RTrim
            cReturn.Add vbCrLf, StrRepeat(vbTab, lngIndentLevel)
            blnNewline = False
            blnAddedNewline = True
        ElseIf cReturn.Length = 0 Then
            ' Avoid adding a newline at beginning
            blnAddedNewline = True
        Else
            blnAddedNewline = False
        End If

        ' Display comments directly where they appear in the source
        If IsTokenType(intTokenType, ttComment, ttBlockComment) Then
            If intTokenType = ttBlockComment Then
                ' Indent multiline block comment to current indent level
                strIndent = StrRepeat(vbTab, lngIndentLevel)
                cReturn.RTrim " " & vbTab
                cReturn.Add vbCrLf, strIndent
                cReturn.Add Replace(strTokenValue, vbCrLf, vbCrLf & strIndent)
            Else
                cReturn.Add strTokenValue
            End If
            blnNewline = True
            GoTo NextToken
        End If

        ' If inside parentheses
        If blnInlineParentheses Then
            ' Check for end of inline parentheses
            If strTokenValue = ")" Then
                cReturn.RTrim

                If blnInlineIndented Then
                    DecreaseIndent colIndents, lngIndentLevel
                    cReturn.Add vbCrLf, StrRepeat(vbTab, lngIndentLevel)
                End If

                blnInlineParentheses = False
                cReturn.Add ") "
                GoTo NextToken
            End If

            ' Break to new line if we reach 30 characters
            If strTokenValue = "," Then
                If intInlineCount >= 30 Then
                    intInlineCount = 0
                    blnNewline = True
                End If
            End If

            ' Keep count of characters within parentheses
            intInlineCount = intInlineCount + Len(strTokenValue)
        End If

        ' Opening parentheses increase the block indent level and start a new line
        If strTokenValue = "(" Then
            ' First check if this should be an inline parentheses block
            ' Examples are "NOW()", "COUNT(*)", "int(10)", key(`somecolumn`), DECIMAL(7,2)
            ' Allow up to 3 non-whitespace tokens inside inline parentheses
            lngLength = 0
            ' Begin secondary loop to look at the next tokens without changing
            ' primary token iteration.
            lngToken2 = lngToken
            Do
                ' Get next non-whitespace token
                lngToken2 = GetNextTokenID(lngToken2, ttWhitespace)

                If lngToken2 > m_colTokens.Count _
                    Or lngToken2 > (lngToken + 250) Then
                    ' Reached end of string
                    Exit Do
                End If

                ' Get type and value of next token
                intNextType = m_colTokens(lngToken2)(0)
                strNextValue = m_colTokens(lngToken2)(1)

                ' Reached closing parentheses, able to inline it
                If strNextValue = ")" Then
                    blnInlineParentheses = True
                    intInlineCount = 0
                    blnInlineIndented = False
                    Exit Do
                End If

                ' Reached an invalid token for inline parentheses
                If strNextValue = ";" Or strNextValue = "(" Then
                    Exit Do
                End If

                ' Reached an invalid token type for inline parentheses
                Select Case intTokenType
                    Case ttReservedTopLevel, _
                        ttReservedNewline, _
                        ttComment, _
                        ttBlockComment
                        Exit Do
                End Select

                ' Add to total length
                lngLength = lngLength + (Len(strNextValue))

            ' Look at next token ahead of current position
            Loop

            If blnInlineParentheses And (lngLength > 30) Then
                blnIncreateBlockIndent = True
                blnInlineIndented = True
                blnNewline = True
            End If

            ' Take out the preceding space unless there was whitespace there in the original query
            If m_colTokens.Count > 0 Then
                If Not m_colTokens(lngToken - 1)(0) = ttWhitespace Then
                    cReturn.RTrim
                End If
            End If

            If Not blnInlineParentheses Then
                blnIncreateBlockIndent = True
                ' Add a newline after the parentheses
                blnNewline = True
            End If

        ElseIf strTokenValue = ")" Then
            ' Closing parentheses decrease the block indent level
            ' Remove whitespace before the closing parentheses
            cReturn.RTrim

            ' Reset indent level
            Do While colIndents.Count > 0
                ' Get type of current indent
                strType = colIndents(1)
                ' Reduce the indents
                DecreaseIndent colIndents, lngIndentLevel
                ' Exit after removing a block indent
                If strType = "block" Then Exit Do
            Loop

            ' Add a newline before the closing parentheses (if not already added)
            If Not blnAddedNewline Then
                cReturn.Add vbCrLf, StrRepeat(vbTab, lngIndentLevel)
            End If

        ElseIf intTokenType = ttReservedTopLevel Then
            ' Top level reserved words start a new line and increase the special indent level
            blnIncreaseSpecialIndent = True

            ' If the last indent type was 'special', decrease the special indent for this round
            If colIndents.Count > 0 Then
                If colIndents(1) = "special" Then
                    DecreaseIndent colIndents, lngIndentLevel
                End If
            End If

            ' Add a newline after the top level reserved word
            blnNewline = True
            ' Add a newline before the top level reserved word (if not already added)
            If Not blnAddedNewline Then
                cReturn.RTrim
                cReturn.Add vbCrLf, StrRepeat(vbTab, lngIndentLevel)
            Else
                ' If we already added a newline, redo the indentation since it may be different now
                cReturn.RTrim vbTab
                cReturn.Add StrRepeat(vbTab, lngIndentLevel)
            End If

            ' Purge any extra whitespace
            PurgeExtraWhitespace strTokenValue

            ' If SQL 'LIMIT' clause, start variable to reset newline
            If (UCase(strTokenValue) = "LIMIT") And Not blnInlineParentheses Then
                blnClauseLimit = True
            End If

        ElseIf blnClauseLimit And (strTokenValue <> ",") _
            And (Not IsTokenType(intTokenType, ttNumber, ttWhitespace)) Then
            ' Checks if we are out of the limit clause
            blnClauseLimit = False

        ElseIf (strTokenValue = ",") And Not blnInlineParentheses Then
            ' Commas start a new line (unless within inline parentheses or SQL 'LIMIT' clause)
            ' If the previous TOKEN_VALUE is 'LIMIT', resets new line
            If blnClauseLimit Then
                blnNewline = False
                blnClauseLimit = False
            Else
                ' All other cases of commas
                blnNewline = True
            End If

        ElseIf intTokenType = ttReservedNewline Then
            ' Newline reserved words start a new line
            ' Add a newline before the reserved word (if not already added)
            If Not blnAddedNewline Then
                cReturn.RTrim
                cReturn.Add vbCrLf, StrRepeat(vbTab, lngIndentLevel)
            End If
            PurgeExtraWhitespace strTokenValue

        ElseIf intTokenType = ttBoundary Then
            ' Multiple boundary characters in a row should not have
            ' spaces between them (not including parentheses)
            lngToken2 = GetPreviousTokenID(lngToken, ttWhitespace)
            If m_colTokens(lngToken2)(0) = ttBoundary Then
                ' If previous non-whitespace character was a boundary, then trim
                ' any whitespace going back to the boundary.
                cReturn.RTrim
            End If
        End If

        ' If the token shouldn't have a space before it
        If strTokenValue = "." _
            Or strTokenValue = "," _
            Or strTokenValue = ";" Then
            ' Trim any whitespace
            cReturn.RTrim
        End If

        ' Add current token value, followed by a space
        cReturn.Add strTokenValue, " "

        ' If the token shouldn't have a space after it
        If strTokenValue = "(" Or strTokenValue = "." Then
            cReturn.RTrim
        End If

        ' If this is the "-" of a negative number, it shouldn't have a space after it
        If strTokenValue <> "-" Then
            GoTo NextToken
        End If

        ' Check next token for number
        lngToken2 = GetNextTokenID(lngToken, ttWhitespace)
        If lngToken2 = 0 Then GoTo NextToken
        If m_colTokens(lngToken2)(0) <> ttNumber Then
            GoTo NextToken
        End If

        ' Check previous token
        lngToken2 = GetPreviousTokenID(lngToken, ttWhitespace)
        If lngToken2 = 0 Then
            GoTo NextToken
        Else
            intTokenType = m_colTokens(lngToken2)(0)
            If IsTokenType(intTokenType, _
                ttQuote, _
                ttBacktickQuote, _
                ttWord, _
                ttNumber) Then
                GoTo NextToken
            End If
        End If

        ' Trim whitespace after dash
        cReturn.RTrim

NextToken:
    Next lngToken

    ' If there are unmatched parentheses
    For Each varItem In colIndents
        If CStr(varItem) = "block" Then
            cReturn.RTrim
            cReturn.Add vbCrLf, "WARNING: unclosed parentheses or section"
            Exit For
        End If
    Next varItem

    ' Trim any trailing spaces from completed string
    cReturn.RTrim

    ' Final formatting of tab indents when returning formatted string
    FormatSQL = Replace(cReturn.GetStr, vbTab, "  ")

    Perf.OperationEnd
    Perf.CategoryEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : Tokenize
' Author    : Adam Waller
' Date      : 8/12/2023
' Purpose   : Take a SQL string and break it down into tokens.
'           : Each token is an array with type(0) and value(1).
'---------------------------------------------------------------------------------------
'
Private Sub Tokenize(strSql As String, intDialect As eSqlDialect)

    Const cstrBreakAfter As String = "LIMIT" & ";;;;"

    Dim strMatch As String
    Dim lngLastPos As Long

    ' Reset collection of token items
    Set m_colTokens = New Collection
    m_strSql = strSql
    m_lngPos = 1
    m_varWordCache(1) = Empty
    m_varWordCache(2) = Empty

    ' Set SQL dialect
    m_intDialect = intDialect

    Perf.CategoryStart "Tokenize SQL"

    ' Loop through SQL, converting string into tokens
    Do While m_lngPos < Len(strSql)

        ' Debugging breakpoint
        If NextChar(cstrBreakAfter) Then Stop

        ' Sequentially move through various possible types of tokens,
        ' adding the token when we find a match.

        ' Whitespace
        If MatchWhitespace(strMatch) Then
            AddToken ttWhitespace, strMatch

        ' Date literal (Microsoft Access SQL)
        ElseIf MatchAccessDateLiteral(strMatch) Then
            AddToken ttQuote, strMatch

        ' Single line comment
        ElseIf NextChar("#") Or NextChar("--") Then
            AddToken ttComment, GetTill(vbCrLf, vbLf)

        ' Block comment
        ElseIf NextChar("/*") Then
            AddToken ttBlockComment, GetTill("*/")

        ' Backtick quote
        ElseIf NextChar("`") Then
            AddToken ttBacktickQuote, GetQuotedString

        ' Quoted string
        ElseIf NextChar("""") Or NextChar("'") Or NextChar("[") Then
            AddToken ttQuote, GetQuotedString

        ' User defined variable
        ElseIf (NextChar("@") Or NextChar(":")) And (RemainingChars > 1) Then

            ' Check for quoted variable name
            If PeekChar(1, """") Or PeekChar(1, "`") Or PeekChar(1, "'") Then
                AddToken ttVariable, GetRange & GetQuotedString(1)
            Else
                ' Non-quoted variable name
                If HasMatches("^(" & GetRange & "[a-zA-Z0-9\._\$]+)", strMatch) Then
                    AddToken ttVariable, strMatch
                Else
                    ' Failed to parse name
                    'TODO: log error
                End If
            End If

        ' Number (decimal, bindary, or hex)
        ElseIf MatchNumber(strMatch) Then
            AddToken ttNumber, strMatch

        ' Boundary character (punctuation and symbols)
        ElseIf NextBoundary = m_lngPos Then
            AddToken ttBoundary, GetRange

        ' A reserved word cannot be preceded by a "."
        ' This makes it so in "mytable.from", "from" is not considered a reserved word
        ElseIf PeekChar(-1, ".") Then

            ' Likely an object name
            If HasMatches("^(.*?)($|\s|[""\'`]|" & RegExBoundaries & ")", strMatch) Then
                AddToken ttWord, strMatch
            Else
                Stop
            End If

        ' Check for reserved words or functions
        Else

            ' Top level reserved word
            If InList(eltReservedToplevel, strMatch) Then
                AddToken ttReservedTopLevel, strMatch

            ' Newline reserved word
            ElseIf InList(eltReservedNewline, strMatch) Then
                AddToken ttReservedNewline, strMatch

            ' Other reserved word
            ElseIf InList(eltReserved, strMatch) Then
                AddToken ttReserved, strMatch

            ' A function must be followed by "("
            ' This makes it so "count(" is considered a function, but "count" alone is not
            ElseIf HasMatches("^((" & cstrFunctions & ")[(]|\s|[)])", strMatch) Then
                ' Add the function, but not the opening parenthesis
                AddToken ttReserved, GetRange(, Len(strMatch) - 1)

            ' Non-reserved word
            ElseIf HasMatches("^(.*?)($|\s|[""\'`]|" & RegExBoundaries & ")", strMatch) Then
                AddToken ttWord, strMatch

            Else
                ' TODO: check for errors
                Stop

            End If

        End If

        ' Make sure we don't get stuck on the same position
        If m_lngPos = lngLastPos Then
            Log.Error eelError, "Unable tp parse SQL after position " & m_lngPos, ModuleName(Me) & ".Tokenize"
            AddToken ttUnknown, GetRange(Len(m_strSql) - m_lngPos)
            Exit Do
        Else
            lngLastPos = m_lngPos
        End If

    ' Move to next token
    Loop

    ' Reset position
    m_lngPos = 1

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetMatches
' Author    : Adam Waller
' Date      : 8/12/2023
' Purpose   : Returns true if we match on the RegEx expression, and sets objMatches to
'           : the resulting match.
'---------------------------------------------------------------------------------------
'
Private Function HasMatches(strRegEx As String, ByRef strMatches As String) As Boolean

    Static oReg As VBScript_RegExp_55.RegExp

    If oReg Is Nothing Then Set oReg = New VBScript_RegExp_55.RegExp

    Dim objMatches As VBScript_RegExp_55.MatchCollection

    Perf.OperationStart "RegEx"
    With oReg ' New VBScript_RegExp_55.RegExp
        .Pattern = strRegEx
        .IgnoreCase = True
        .Global = False
        Set objMatches = .Execute(Mid$(m_strSql, m_lngPos))
        If objMatches.Count = 0 Then
            strMatches = vbNullString
        Else
            If objMatches(0).SubMatches.Count > 0 Then
                ' Use first submatch, if found
                If Len(objMatches(0).SubMatches(0)) Then
                    strMatches = objMatches(0).SubMatches(0)
                Else
                    ' Use primary match
                    strMatches = objMatches(0)
                End If
            Else
                ' Fall back to primary match
                strMatches = objMatches(0)
            End If
            HasMatches = True
        End If
    End With
    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : AddToken
' Author    : Adam Waller
' Date      : 8/12/2023
' Purpose   : Add a token to the current collection and moves the position marker
'---------------------------------------------------------------------------------------
'
Private Sub AddToken(intType As eTokenTypes, strValue As String)
    Perf.OperationStart "Add Token"
    m_colTokens.Add Array(intType, strValue)
    m_lngPos = m_lngPos + Len(strValue)
    Perf.OperationEnd
End Sub


'---------------------------------------------------------------------------------------
' Procedure : NextChar
' Author    : Adam Waller
' Date      : 8/12/2023
' Purpose   : Return next one or more characters from current position.
'---------------------------------------------------------------------------------------
'
Private Function NextChar(strMatch As String) As Boolean
    If (m_lngPos + 1) > Len(m_strSql) Then Exit Function
    NextChar = (Mid$(m_strSql, m_lngPos, Len(strMatch)) = strMatch)
End Function


'---------------------------------------------------------------------------------------
' Procedure : PeekChar
' Author    : Adam Waller
' Date      : 8/12/2023
' Purpose   : Peek ahead at another character other than the next one.
'---------------------------------------------------------------------------------------
'
Private Function PeekChar(lngOffset As Long, strMatchChar As String) As Boolean
    If (m_lngPos + lngOffset) < 1 Then Exit Function
    If (m_lngPos + lngOffset) > Len(m_strSql) Then Exit Function
    Perf.OperationStart "Peek Character"
    PeekChar = (Mid$(m_strSql, m_lngPos + lngOffset, 1) = strMatchChar)
    Perf.OperationEnd
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTill
' Author    : Adam Waller
' Date      : 8/12/2023
' Purpose   : Return the string until we find the stop characters, or the remaining
'           : string if the stop characters were not found.
'---------------------------------------------------------------------------------------
'
Private Function GetTill(strStopAt As String, Optional strAltStopAt As String, Optional blnIncludeStopChars As Boolean = False) As String

    Dim lngPos As Long
    Dim lngLen As Long
    Dim strStop As String

    Perf.OperationStart "GetTill"

    ' See if the searched string exists
    lngPos = InStr(m_lngPos + 1, m_strSql, strStopAt)

    ' Try again using alternate stop
    If lngPos = 0 Then
        lngPos = InStr(m_lngPos, m_strSql, strAltStopAt)
        strStop = strAltStopAt
    Else
        strStop = strStopAt
    End If

    ' Make sure we found a match
    If lngPos = 0 Then
        ' No match. Return the remaining string
        GetTill = Mid$(m_strSql, m_lngPos)
    Else
        lngLen = lngPos - m_lngPos
        ' Add in the length of the matched stop characters, if requested
        If blnIncludeStopChars Then lngLen = lngLen + Len(strStop)
        ' Return the string
        GetTill = Mid$(m_strSql, m_lngPos, lngLen)
    End If
    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetRange
' Author    : Adam Waller
' Date      : 8/12/2023
' Purpose   : Get a range of characters (kind of like a relative Mid$() function)
'---------------------------------------------------------------------------------------
'
Private Function GetRange(Optional lngStartOffset As Long, Optional lngLength As Long = 1) As String
    GetRange = Mid$(m_strSql, m_lngPos + lngStartOffset, lngLength)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetQuotedString
' Author    : Adam Waller
' Date      : 8/12/2023
' Purpose   : Return a quoted string (dialect-specific)
'           : https://stackoverflow.com/q/10573922/4121863
'           : https://stackoverflow.com/q/9719869/4121863
'---------------------------------------------------------------------------------------
'
Private Function GetQuotedString(Optional lngStartOffset As Long = 0) As String

    Dim strExp As String
    Dim objMatches As VBScript_RegExp_55.MatchCollection

    Perf.OperationStart "Get Quoted String"
    With New clsConcat
        ' Build out RegEx expression
        .Add "^("

            ' Accomodate dialect-specific variants
            Select Case m_intDialect

                Case esdMySQL
                    ' Backtick quoted string using `` to escape
                    .Add "((`[^`]*($|`))+)|"

                    ' Double quoted string using "" or \" to escape
                    .Add "((""[^""\\\\]*(?:\\\\.[^""\\\\]*)*(""|$))+)|"

                    ' Single quoted string using '' or \' to escape
                    .Add "((\'[^\'\\\\]*(?:\\\\.[^\'\\\\]*)*(\'|$))+)"

                Case Else
                    ' Square bracket quoted string (SQL Server) using ]] to escape
                    .Add "((\[[^\]]*($|\]))(\][^\]]*($|\]))*)|"

                    ' Double quoted string using "" to escape
                    .Add "((""[^""]*(""|$))+)|"

                    ' Single quoted string using '' to escape
                    .Add "((\'[^\']*(\'|$))+)"

            End Select
        .Add ")"
        strExp = .GetStr
    End With

    ' Apply RegEx
    With New VBScript_RegExp_55.RegExp
        .Global = True
        .Multiline = True
        .Pattern = strExp
        Set objMatches = .Execute(Mid$(m_strSql, m_lngPos + lngStartOffset))
        If objMatches.Count > 0 Then GetQuotedString = objMatches(0)
    End With
    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : RemainingChars
' Author    : Adam Waller
' Date      : 8/12/2023
' Purpose   : Return remaining number of characters from current position
'---------------------------------------------------------------------------------------
'
Private Function RemainingChars() As Long
    RemainingChars = Len(m_strSql) - m_lngPos
End Function


'---------------------------------------------------------------------------------------
' Procedure : StrRepeat
' Author    : Adam Waller
' Date      : 8/14/2023
' Purpose   : Repeat a string a specified number of times. (High performance version)
'---------------------------------------------------------------------------------------
'
Private Function StrRepeat(strText As String, lngTimes As Long) As String

    Dim lngCnt As Long
    Dim lngStart As Long
    Dim strReturn As String

    ' Build full length string to avoid slow string contatenation
    strReturn = Space(Len(strText) * lngTimes)

    ' Apply repeated text to return string
    For lngCnt = 1 To lngTimes
        ' Calculate start position
        lngStart = 1 + ((lngCnt * Len(strText)) - Len(strText))
        Mid$(strReturn, lngStart, Len(strText)) = strText
    Next lngCnt

    ' Return repeated string
    StrRepeat = strReturn

End Function


'---------------------------------------------------------------------------------------
' Procedure : InList
' Author    : Adam Waller
' Date      : 8/11/2023
' Purpose   : Returns true if the next word or two are found in the list
'---------------------------------------------------------------------------------------
'
Private Function InList(intList As eListType, ByRef strMatch As String) As Boolean

    Dim intWords As Integer
    Dim intMaxWords As Integer
    Dim strList As String
    Dim strTest As String
    Dim strWords As String
    Dim lngEndPos As Long

    ' Look up list of words
    Select Case intList
        Case eltBoundaries:         strList = cstrBoundaries
        Case eltRegExSpecial:       strList = cstrRegExSpecial
        Case eltFunctions:          strList = cstrFunctions
        Case eltReserved:           strList = cstrReserved
        Case eltReservedNewline:    strList = cstrReservedNewline
        Case eltReservedToplevel:   strList = cstrReservedToplevel
    End Select

    ' Determine max words to look up (i.e. "LEFT JOIN")
    If InStr(1, strList, " ") > 0 Then
        intMaxWords = 2
    Else
        intMaxWords = 1
    End If

    Perf.OperationStart "InList"
    ' Loop to check multiple or single words
    For intWords = intMaxWords To 1 Step -1

        ' Build test string of word(s)
        strWords = GetNextWords(intWords, lngEndPos)
        strTest = "|" & UCase(strWords) & "|"

        ' Make sure we found some words
        If strTest = "||" Then Exit For

        ' See if the words exist in the list
        If InStr(1, strList, strTest) > 0 Then

            ' Found a match. Return original string (including any extra whitespace)
            strMatch = Mid$(m_strSql, m_lngPos, lngEndPos - m_lngPos)
            InList = True
            Exit For
        End If

    Next intWords
    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetNextWords
' Author    : Adam Waller
' Date      : 8/17/2023
' Purpose   : Return the next x words before the next boundary character. (Replacing
'           : any number of whitespace characters with a single space.)
'           : Also returns the ending position in the original string.
'---------------------------------------------------------------------------------------
'
Private Function GetNextWords(intCount As Integer, Optional ByRef lngEndPos As Long) As String

    Dim lngChar As Integer
    Dim lngStart As Long
    Dim intFoundWords As Integer
    Dim lngBoundary As Long

    ' Cache word lookup so we can avoid repeated calls when checking the same word(s)
    ' against different lists. (For 1 or 2 words)
    If intCount = 1 Or intCount = 2 Then
        If IsArray(m_varWordCache(intCount)) Then
            If m_varWordCache(intCount)(0) = m_lngPos Then
                ' Found matching cache from last lookup
                lngEndPos = m_varWordCache(intCount)(1)
                GetNextWords = m_varWordCache(intCount)(2)
                Exit Function
            End If
        End If
    End If

    ' Get next boundary, and exit if we don't have anything before the boundary
    lngBoundary = NextBoundary
    If lngBoundary = m_lngPos Then Exit Function

    Perf.OperationStart "Get Next Words"
    With New clsConcat
        ' Loop through characters up till next boundary
        For lngChar = m_lngPos To lngBoundary - 1
            Select Case AscW(Mid$(m_strSql, lngChar, 1))
                Case 32, 9, 10, 13, 0
                    ' Found whitespace
                    If lngStart > 0 Then
                        ' Complete previous word
                        If intFoundWords > 0 Then .Add " "
                        intFoundWords = intFoundWords + 1
                        .Add Mid$(m_strSql, lngStart, lngChar - lngStart)
                        lngEndPos = lngChar
                        lngStart = 0
                        If intFoundWords >= intCount Then Exit For
                    End If

                Case Else
                    ' Begin at first non-whitespace character
                    If lngStart = 0 Then lngStart = lngChar
                    ' Exit if we reached the boundary
                    If lngChar = lngBoundary - 1 Then
                        .Add Mid$(m_strSql, lngStart, lngChar - lngStart)
                        lngEndPos = lngChar - 1
                        Exit For
                    End If

            End Select
        Next lngChar

        If lngEndPos = 0 Then Stop

        ' Prepare results
        GetNextWords = .GetStr

        ' Save to cache for 1 and 2 word lookups
        If intCount = 1 Or intCount = 2 Then
            m_varWordCache(intCount) = Array(m_lngPos, lngEndPos, GetNextWords)
        End If
    End With
    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : NextBoundary
' Author    : Adam Waller
' Date      : 8/17/2023
' Purpose   : Return the position of the next boundary character
'---------------------------------------------------------------------------------------
'
Private Function NextBoundary() As Long

    Dim lngChar As Long
    Dim lngLowest As Long

    Perf.OperationStart "Next Boundary"
    ' Begin with maximum length, and reduce based on found characters
    lngLowest = Len(m_strSql)

    ' Loop through next characters
    For lngChar = m_lngPos To Len(m_strSql)

        ' Use AscW for fastest performance
        Select Case AscW(Mid$(m_strSql, lngChar, 1))

            ' Boundary characters (Codes extracted from cstrBoundaries)
            Case 44, 59, 58, 41, 40, 46, 61, 60, 62, 43, 45, 42, 47, 33, 94, 37, 124, 38, 35
                NextBoundary = lngChar
                Exit For

            Case Else
                ' Continue to next character

        End Select
    Next lngChar

    ' Return position of first boundary, or length of string if no boundary was found.
    If NextBoundary = 0 Then NextBoundary = Len(m_strSql)

    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : MatchWhitespace
' Author    : Adam Waller
' Date      : 8/18/2023
' Purpose   : Get any contiguous whitespace from the current position.
'---------------------------------------------------------------------------------------
'
Private Function MatchWhitespace(ByRef strMatches As String) As Boolean

    Dim lngChar As Long
    Dim lngFound As Long

    Perf.OperationStart "Get Whitespace"

    ' Loop through next characters
    For lngChar = m_lngPos To Len(m_strSql)

        ' Use AscW for fastest performance
        Select Case AscW(Mid$(m_strSql, lngChar, 1))

            ' Any whitespace character (space, tab, lf, cr, nullchar)
            Case 32, 9, 10, 13, 0
                lngFound = lngChar

            ' Non-whitespace character
            Case Else
                Exit For

        End Select
    Next lngChar

    ' Set match string if we found whitespace
    If lngFound > 0 Then
        strMatches = Mid$(m_strSql, m_lngPos, (lngFound + 1) - m_lngPos)
        MatchWhitespace = True
    End If

    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : MatchAccessDateLiteral
' Author    : Adam Waller
' Date      : 9/21/2023
' Purpose   : Returns a matching date literal for Microsoft Access. (I.e. #1/1/2000#)
'           : Checks for a closing hash, and a date value inside.
'---------------------------------------------------------------------------------------
'
Private Function MatchAccessDateLiteral(ByRef strMatches As String) As Boolean

    Dim lngPos As Long
    Dim lngLen As Long

    ' See if we start with a hash character...
    If AscW(Mid$(m_strSql, m_lngPos, 1)) = 35 Then

        ' Check for closing hash character
        lngPos = InStr(m_lngPos + 1, m_strSql, "#")
        If lngPos > m_lngPos Then
            ' We found another hash somewhere down the line.
            ' See if the distance looks reasonable.
            lngLen = lngPos - m_lngPos
            If lngLen < 50 Then
                ' See if this is a date value
                If IsDate(Mid$(m_strSql, m_lngPos + 1, lngLen - 1)) Then
                    ' Appears to be a valid date. Let's go with it.
                    MatchAccessDateLiteral = True
                    strMatches = Mid$(m_strSql, m_lngPos, lngLen + 1)
                End If
            End If
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : MatchNumber
' Author    : Adam Waller
' Date      : 8/18/2023
' Purpose   : Return a number token from the current position, if found.
'           : Checks for a leading digit before the slower RegEx for the various
'           : types of numbers.
'---------------------------------------------------------------------------------------
'
Private Function MatchNumber(ByRef strMatches As String) As Boolean

    ' Perform an inital very fast test to see if this might be a number.
    Select Case AscW(Mid$(m_strSql, m_lngPos, 1))
        ' Any number 0 to 9
        Case 48 To 57
            ' Continue with RegEx
        Case Else
            ' Not a number
            Exit Function
    End Select

    ' Now move on with the RegEx to get the entire number
    MatchNumber = HasMatches("^([0-9]+(\.[0-9]+)?|0x[0-9a-fA-F]+|0b[01]+)($|\s|""\'`|" & RegExBoundaries & ")", strMatches)

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsTokenType
' Author    : Adam Waller
' Date      : 8/15/2023
' Purpose   : Returns true if the token type matches any of the specified types.
'---------------------------------------------------------------------------------------
'
Private Function IsTokenType(intToken As eTokenTypes, ParamArray intMatchTypes() As Variant) As Boolean
    Dim intMatch As Integer
    For intMatch = 0 To UBound(intMatchTypes)
        If intToken = intMatchTypes(intMatch) Then
            IsTokenType = True
            Exit For
        End If
    Next intMatch
End Function


'---------------------------------------------------------------------------------------
' Procedure : IncreaseIndent
' Author    : Adam Waller
' Date      : 8/15/2023
' Purpose   : Increase the indent level
'---------------------------------------------------------------------------------------
'
Private Sub IncreaseIndent(ByRef colIndents As Collection, ByRef lngIndentLevel As Long, strIndentType As String)
    If colIndents Is Nothing Then Exit Sub
    If colIndents.Count = 0 Then
        colIndents.Add strIndentType
    Else
        colIndents.Add strIndentType, , 1
    End If
    lngIndentLevel = lngIndentLevel + 1
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DecreaseIndent
' Author    : Adam Waller
' Date      : 8/15/2023
' Purpose   : Remove the first element of the indents collection, shifting the remaining
'           : elements down. (Similar to the PHP array_shift function)
'           : Also reduce the indent level by one.
'---------------------------------------------------------------------------------------
'
Private Sub DecreaseIndent(ByRef colIndents As Collection, ByRef lngIndentLevel As Long)
    If colIndents Is Nothing Then Exit Sub
    If colIndents.Count = 0 Then Exit Sub
    colIndents.Remove 1
    lngIndentLevel = lngIndentLevel - 1
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PurgeExtraWhitespace
' Author    : Adam Waller
' Date      : 8/15/2023
' Purpose   : Replace any extra whitespace within the string with a single space.
'---------------------------------------------------------------------------------------
'
Private Sub PurgeExtraWhitespace(ByRef strValue As String)

    Dim strResult As String

    ' Check for any space, tab, or vbcrlf
    If (InStr(1, strValue, " ") > 0) _
        Or (InStr(1, strValue, vbTab) > 0) _
        Or (InStr(1, strValue, vbCrLf) > 0) Then

        With New VBScript_RegExp_55.RegExp
            .Multiline = True
            '.Global = True
            .Pattern = "\s+"
            strResult = .Replace(strValue, " ")
            If Len(strResult) Then
                If strResult <> strValue Then strValue = strResult
            End If
        End With
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetPreviousToken
' Author    : Adam Waller
' Date      : 8/15/2023
' Purpose   : Return the previous token ID, optionally skipping any excluded type.
'---------------------------------------------------------------------------------------
'
Private Function GetPreviousTokenID(lngCurrentToken As Long, Optional intExceptType As eTokenTypes) As Long
    Dim intToken As Integer
    For intToken = lngCurrentToken - 1 To 1 Step -1
        If intToken = 0 Then
            ' Could not find a matching token, or no previous token
            GetPreviousTokenID = 0
        End If
        If m_colTokens(intToken)(0) <> intExceptType Then
            GetPreviousTokenID = intToken
            Exit For
        End If
    Next intToken
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetNextTokenID
' Author    : Adam Waller
' Date      : 8/15/2023
' Purpose   : Return the next token ID, optionally skipping any excluded type.
'---------------------------------------------------------------------------------------
'
Private Function GetNextTokenID(lngCurrentToken As Long, Optional intExceptType As eTokenTypes) As Long
    Dim intToken As Integer

    If lngCurrentToken + 1 > m_colTokens.Count Then
        GetNextTokenID = lngCurrentToken + 1
        Exit Function
    End If

    For intToken = lngCurrentToken + 1 To m_colTokens.Count
        If m_colTokens(intToken)(0) <> intExceptType Then
            GetNextTokenID = intToken
            Exit For
        End If
    Next intToken
End Function


'---------------------------------------------------------------------------------------
' Procedure : RegExBoundaries
' Author    : Adam Waller
' Date      : 8/12/2023
' Purpose   : Return a string of properly escaped regex boundary characters.
'---------------------------------------------------------------------------------------
'
Private Function RegExBoundaries() As String

    Dim lngPos As Long
    Dim strChar As String

    Static strBoundaries As String

    ' Check for cached list
    If strBoundaries = vbNullString Then

        ' Build out
        With New clsConcat
            .Add "("
            .AppendOnAdd = "|"
            For lngPos = 1 To Len(cstrBoundaries)
                strChar = Mid$(cstrBoundaries, lngPos, 1)
                If InStr(1, cstrRegExSpecial, strChar) > 0 Then
                    ' Escape this character
                    .Add "\", strChar
                Else
                    .Add strChar
                End If
            Next lngPos
            ' Trim final delimiter and cache result
            .Remove 1
            strBoundaries = .GetStr & ")"
        End With
    End If

    ' Return list of boundaries
    RegExBoundaries = strBoundaries

End Function


'#######################################################################################
'#######################################################################################
'
'                                   TESTING CODE
'
'   The following section contains code and functions used to verify the behavior
'   of this class with various types of queries to ensure consistent output.
'
'#######################################################################################
'#######################################################################################


'---------------------------------------------------------------------------------------
' Procedure : SelfTest
' Author    : Adam Waller
' Date      : 8/16/2023
' Purpose   : Perform testing with various sample queries to ensure that the formatting
'           : rules are being applied as intended.
'           : More complex tests can be performed using larger queries and result
'           : files, but this should allow us to cover the primary functionality.
'---------------------------------------------------------------------------------------
'
Public Sub SelfTest()

    Dim strActual As String

    ' Test GetNextWords
    Tokenize "  LEFT " & vbTab & vbCrLf & " JOIN test on 1=2", esdAccess
    Debug.Assert GetNextWords(2) = "LEFT JOIN"
    Debug.Assert GetNextWords(1) = "LEFT"

    ' Test simple query with a few features
    Tokenize "SELECT  5 AS `TEST`", esdMySQL

    ' Verify tokens
    Debug.Assert m_colTokens.Count = 7
    Debug.Assert VerifyToken(1, ttReservedTopLevel, "SELECT")
    Debug.Assert VerifyToken(2, ttWhitespace, "  ")
    Debug.Assert VerifyToken(3, ttNumber, "5")
    Debug.Assert VerifyToken(4, ttWhitespace, " ")
    Debug.Assert VerifyToken(5, ttReserved, "AS")
    Debug.Assert VerifyToken(6, ttWhitespace, " ")
    Debug.Assert VerifyToken(7, ttBacktickQuote, "`TEST`")

    ' Verify result
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  5 AS `TEST`"
        .Remove 2
        strActual = .GetStr
    End With
    Debug.Assert (strActual = FormatSQL)
    If (strActual <> FormatSQL) Then Diff.Strings strActual, FormatSQL


    ' Test Access date literal with MySQL inline comment
    Tokenize "SELECT (#1/1/2000#) AS SampleDate # MySQL inline ## comment", esdAccess

    ' Verify tokens
    Debug.Assert m_colTokens.Count = 11
    Debug.Assert VerifyToken(1, ttReservedTopLevel, "SELECT")
    Debug.Assert VerifyToken(2, ttWhitespace, " ")
    Debug.Assert VerifyToken(3, ttBoundary, "(")
    Debug.Assert VerifyToken(4, ttQuote, "#1/1/2000#")
    Debug.Assert VerifyToken(5, ttBoundary, ")")
    Debug.Assert VerifyToken(6, ttWhitespace, " ")
    Debug.Assert VerifyToken(7, ttReserved, "AS")
    Debug.Assert VerifyToken(8, ttWhitespace, " ")
    Debug.Assert VerifyToken(9, ttWord, "SampleDate")
    Debug.Assert VerifyToken(10, ttWhitespace, " ")
    Debug.Assert VerifyToken(11, ttComment, "# MySQL inline ## comment")

    ' Verify result
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  (#1/1/2000#) AS SampleDate # MySQL inline ## comment"
        .Remove 2
        strActual = .GetStr
    End With
    Debug.Assert (strActual = FormatSQL)
    If (strActual <> FormatSQL) Then Diff.Strings strActual, FormatSQL


    ' Example query from https://github.com/doctrine/sql-formatter
    Tokenize "SELECT count(*),`Column1`,`Testing`, `Testing Three` FROM `Table1`" & _
            " WHERE Column1 = 'testing' AND ( (`Column2` = `Column3` OR Column4 >= NOW()) )" & _
            " GROUP BY Column1 ORDER BY Column3 DESC LIMIT 5,10", esdMySQL

    ' Verify tokens
    Debug.Assert m_colTokens.Count = 66
    Debug.Assert VerifyToken(1, ttReservedTopLevel, "SELECT")
    Debug.Assert VerifyToken(2, ttWhitespace, " ")
    Debug.Assert VerifyToken(3, ttReserved, "count")
    Debug.Assert VerifyToken(4, ttBoundary, "(")
    Debug.Assert VerifyToken(5, ttBoundary, "*")
    Debug.Assert VerifyToken(6, ttBoundary, ")")
    Debug.Assert VerifyToken(7, ttBoundary, ",")
    Debug.Assert VerifyToken(8, ttBacktickQuote, "`Column1`")
    Debug.Assert VerifyToken(9, ttBoundary, ",")
    Debug.Assert VerifyToken(10, ttBacktickQuote, "`Testing`")
    Debug.Assert VerifyToken(11, ttBoundary, ",")
    Debug.Assert VerifyToken(12, ttWhitespace, " ")
    Debug.Assert VerifyToken(13, ttBacktickQuote, "`Testing Three`")
    Debug.Assert VerifyToken(14, ttWhitespace, " ")
    Debug.Assert VerifyToken(15, ttReservedTopLevel, "FROM")
    Debug.Assert VerifyToken(16, ttWhitespace, " ")
    Debug.Assert VerifyToken(17, ttBacktickQuote, "`Table1`")
    Debug.Assert VerifyToken(18, ttWhitespace, " ")
    Debug.Assert VerifyToken(19, ttReservedTopLevel, "WHERE")
    Debug.Assert VerifyToken(20, ttWhitespace, " ")
    Debug.Assert VerifyToken(21, ttWord, "Column1")
    Debug.Assert VerifyToken(22, ttWhitespace, " ")
    Debug.Assert VerifyToken(23, ttBoundary, "=")
    Debug.Assert VerifyToken(24, ttWhitespace, " ")
    Debug.Assert VerifyToken(25, ttQuote, "'testing'")
    Debug.Assert VerifyToken(26, ttWhitespace, " ")
    Debug.Assert VerifyToken(27, ttReservedNewline, "AND")
    Debug.Assert VerifyToken(28, ttWhitespace, " ")
    Debug.Assert VerifyToken(29, ttBoundary, "(")
    Debug.Assert VerifyToken(30, ttWhitespace, " ")
    Debug.Assert VerifyToken(31, ttBoundary, "(")
    Debug.Assert VerifyToken(32, ttBacktickQuote, "`Column2`")
    Debug.Assert VerifyToken(33, ttWhitespace, " ")
    Debug.Assert VerifyToken(34, ttBoundary, "=")
    Debug.Assert VerifyToken(35, ttWhitespace, " ")
    Debug.Assert VerifyToken(36, ttBacktickQuote, "`Column3`")
    Debug.Assert VerifyToken(37, ttWhitespace, " ")
    Debug.Assert VerifyToken(38, ttReservedNewline, "OR")
    Debug.Assert VerifyToken(39, ttWhitespace, " ")
    Debug.Assert VerifyToken(40, ttWord, "Column4")
    Debug.Assert VerifyToken(41, ttWhitespace, " ")
    Debug.Assert VerifyToken(42, ttBoundary, ">")
    Debug.Assert VerifyToken(43, ttBoundary, "=")
    Debug.Assert VerifyToken(44, ttWhitespace, " ")
    Debug.Assert VerifyToken(45, ttWord, "NOW")
    Debug.Assert VerifyToken(46, ttBoundary, "(")
    Debug.Assert VerifyToken(47, ttBoundary, ")")
    Debug.Assert VerifyToken(48, ttBoundary, ")")
    Debug.Assert VerifyToken(49, ttWhitespace, " ")
    Debug.Assert VerifyToken(50, ttBoundary, ")")
    Debug.Assert VerifyToken(51, ttWhitespace, " ")
    Debug.Assert VerifyToken(52, ttReservedTopLevel, "GROUP BY")
    Debug.Assert VerifyToken(53, ttWhitespace, " ")
    Debug.Assert VerifyToken(54, ttWord, "Column1")
    Debug.Assert VerifyToken(55, ttWhitespace, " ")
    Debug.Assert VerifyToken(56, ttReservedTopLevel, "ORDER BY")
    Debug.Assert VerifyToken(57, ttWhitespace, " ")
    Debug.Assert VerifyToken(58, ttWord, "Column3")
    Debug.Assert VerifyToken(59, ttWhitespace, " ")
    Debug.Assert VerifyToken(60, ttReserved, "DESC")
    Debug.Assert VerifyToken(61, ttWhitespace, " ")
    Debug.Assert VerifyToken(62, ttReservedTopLevel, "LIMIT")
    Debug.Assert VerifyToken(63, ttWhitespace, " ")
    Debug.Assert VerifyToken(64, ttNumber, "5")
    Debug.Assert VerifyToken(65, ttBoundary, ",")
    Debug.Assert VerifyToken(66, ttNumber, "10")

    ' Verify result
    With New clsConcat
        .AppendOnAdd = vbCrLf
        .Add "SELECT"
        .Add "  count(*),"
        .Add "  `Column1`,"
        .Add "  `Testing`,"
        .Add "  `Testing Three`"
        .Add "FROM"
        .Add "  `Table1`"
        .Add "WHERE"
        .Add "  Column1 = 'testing'"
        .Add "  AND ("
        .Add "    ("
        .Add "      `Column2` = `Column3`"
        .Add "      OR Column4 >= NOW()"
        .Add "    )"
        .Add "  )"
        .Add "GROUP BY"
        .Add "  Column1"
        .Add "ORDER BY"
        .Add "  Column3 DESC"
        .Add "LIMIT"
        .Add "  5, 10"
        .Remove 2
        strActual = .GetStr
    End With
    Debug.Assert (strActual = FormatSQL)
    If (strActual <> FormatSQL) Then Diff.Strings strActual, FormatSQL


    'PrintTokens
    'BuildTestFromTokens
    'Diff.Strings strActual, FormatSQL

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Perf
' Author    : Adam Waller
' Date      : 8/17/2023
' Purpose   : Subclass global performance function to turn on or off for this class
'           : independent of the global performance monitoring.
'---------------------------------------------------------------------------------------
'
Private Function Perf() As clsPerformance
    Static cInternal As clsPerformance
    If cInternal Is Nothing Then
        Set cInternal = New clsPerformance
        ' Disable detailed internal performance monitoring by default
        cInternal.Enabled = False
    End If
    Set Perf = cInternal
End Function


'---------------------------------------------------------------------------------------
' Procedure : VerifyToken
' Author    : Adam Waller
' Date      : 8/16/2023
' Purpose   : Verify a token type and value based on index position.
'---------------------------------------------------------------------------------------
'
Private Function VerifyToken(intIndex As Integer, intType As eTokenTypes, strValue As String) As Boolean
    If m_colTokens Is Nothing Then Exit Function
    If intIndex > m_colTokens.Count Then Exit Function
    If m_colTokens(intIndex)(0) = intType Then
        VerifyToken = (m_colTokens(intIndex)(1) = strValue)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildTestFromTokens
' Author    : Adam Waller
' Date      : 8/16/2023
' Purpose   : Builds a verification test from the current tokens.
'---------------------------------------------------------------------------------------
'
Private Function BuildTestFromTokens()

    Dim intToken As Integer
    Dim strType As String
    Dim strValue As String
    Dim varLines As Variant
    Dim intLine As Integer

    Debug.Print vbCrLf & vbCrLf & "    ' Verify tokens"
    Debug.Print "    Debug.Assert m_colTokens.Count = " & m_colTokens.Count

    ' Loop through tokens
    For intToken = 1 To m_colTokens.Count

        ' Get type and value
        strType = TypeEnumToString(m_colTokens(intToken)(0))
        strValue = Replace(m_colTokens(intToken)(1), """", """""")

        ' Print test to debug window
        Debug.Print "    Debug.Assert VerifyToken(" & intToken & ", " & strType & ", """ & strValue & """)"

    Next intToken

    ' Output formatted SQL
    varLines = Split(FormatSQL, vbCrLf)
    Debug.Print vbCrLf & "    ' Verify result"
    Debug.Print "    With New clsConcat"
    Debug.Print "        .AppendOnAdd = vbCrLf"

    ' Loop through lines, outputting SQL with line breaks
    For intLine = 0 To UBound(varLines)
        Debug.Print "        .Add """ & Replace(varLines(intLine), """", """""") & """"
    Next intLine
    Debug.Print "        .Remove 2"
    Debug.Print "        strActual = .GetStr"
    Debug.Print "    End With"
    Debug.Print "    Debug.Assert (strActual = FormatSQL)"
    Debug.Print "    If (strActual <> FormatSQL) Then Diff.Strings strActual, FormatSQL"

End Function


'---------------------------------------------------------------------------------------
' Procedure : PrintTokens
' Author    : Adam Waller
' Date      : 8/14/2023
' Purpose   : Print the current sequence of tokens (type and value) for debugging
'---------------------------------------------------------------------------------------
'
Private Sub PrintTokens()

    Dim varToken As Variant
    Dim strType As String

    If m_colTokens Is Nothing Then
        Debug.Print "No tokens found"
        Exit Sub
    End If

    ' Print header
    Debug.Print "---------------------------------------------------"
    Debug.Print "TOKEN NAME          TOKEN VALUE"
    Debug.Print "---------------------------------------------------"

    ' Loop through Tokens
    For Each varToken In m_colTokens

        ' Print type name with padding to start values at same position
        strType = TypeEnumToString(varToken(0))
        Debug.Print strType;

        ' Add padding to start values at same position
        Debug.Print Space(20 - Len(strType));

        ' Print value
        Debug.Print varToken(1)

    Next varToken
    Debug.Print "---------------------------------------------------"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TypeEnumToString
' Author    : Adam Waller
' Date      : 8/16/2023
' Purpose   : Convert token type enum to string. (Used for debug output and automated
'           : test building.)
'---------------------------------------------------------------------------------------
'
Private Function TypeEnumToString(intType As Variant) As String
    Select Case intType
        Case ttUnknown:             TypeEnumToString = "ttUnknown"
        Case ttWhitespace:          TypeEnumToString = "ttWhitespace"
        Case ttWord:                TypeEnumToString = "ttWord"
        Case ttQuote:               TypeEnumToString = "ttQuote"
        Case ttBacktickQuote:       TypeEnumToString = "ttBacktickQuote"
        Case ttReserved:            TypeEnumToString = "ttReserved"
        Case ttReservedTopLevel:    TypeEnumToString = "ttReservedTopLevel"
        Case ttReservedNewline:     TypeEnumToString = "ttReservedNewline"
        Case ttBoundary:            TypeEnumToString = "ttBoundary"
        Case ttComment:             TypeEnumToString = "ttComment"
        Case ttBlockComment:        TypeEnumToString = "ttBlockComment"
        Case ttNumber:              TypeEnumToString = "ttNumber"
        Case ttError:               TypeEnumToString = "ttError"
        Case ttVariable:            TypeEnumToString = "ttVariable"
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAscWFromString
' Author    : Adam Waller
' Date      : 8/18/2023
' Purpose   : Print out AscW indexes of each character in the string.
'           : (Use this when building case statements)
'---------------------------------------------------------------------------------------
'
Private Function GetAscWFromString(strText As String) As String
    Dim lngCnt As Long
    Debug.Print
    For lngCnt = 1 To Len(strText)
        Debug.Print AscW(Mid$(strText, lngCnt, 1)) & ", ";
    Next lngCnt
End Function


'---------------------------------------------------------------------------------------
' Procedure : PerformanceTesting
' Author    : Adam Waller
' Date      : 8/18/2023
' Purpose   : Test performance of tokenizing and building formatted string
'---------------------------------------------------------------------------------------
'
Private Function TestPerformance()

    Dim lngCnt As Long
    Dim lngMax As Long

    ' Number of iterations to perform
    lngMax = 100

    ' Test performance of tokenizing a query
    Perf.Enabled = True
    Perf.StartTiming
    For lngCnt = 1 To lngMax
        Tokenize "SELECT count(*),`Column1`,`Testing`, `Testing Three` FROM `Table1`" & _
                " WHERE Column1 = 'testing' AND ( (`Column2` = `Column3` OR Column4 >= NOW()) )" & _
                " GROUP BY Column1 ORDER BY Column3 DESC LIMIT 5,10", esdMySQL
    Next lngCnt

    ' Test performance of formatting SQL
    For lngCnt = 1 To lngMax
        FormatSQL
    Next lngCnt

    Perf.EndTiming
    Debug.Print Perf.GetReports
    Perf.Enabled = False

End Function
