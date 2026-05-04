Attribute VB_Name = "modConnect"
'---------------------------------------------------------------------------------------
' Module    : modConnect
' Author    : hecon5
' Date      : 03/16/2021
' Purpose   : Functions to deal with connection strings for tables, queries, etc.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Utility")

Private Const ModuleName = "modConnect"
Private Const ENV_KEY_PREFIX As String = "conn_"
Private m_dCachedConnections As Dictionary
Private m_dBackEndConnections As Dictionary
Private m_dUnavailableBackEnds As Dictionary
Private m_cEnvCache As clsDotEnv
Private m_dEnvKeysWritten As Dictionary
Private m_dMissingEnvKeys As Dictionary


'---------------------------------------------------------------------------------------
' Procedure : SanitizeConnectionString
' Author    : Adam Waller
' Date      : 2/26/2021
' Purpose   : Sanitize the connection string by removing unneeded information and
'           : converting database path to relative.
'---------------------------------------------------------------------------------------
'
Public Function SanitizeConnectionString(strConnection As String) As String

    Dim lngPart As Long
    Dim varParts As Variant
    Dim strPart As String

    If strConnection = vbNullString Then Exit Function

    ' Create array of connection string parts
    varParts = Split(strConnection, ";")

    ' Loop through parts, building new connection string
    With New clsConcat
        .AppendOnAdd = ";"
        For lngPart = 0 To UBound(varParts)
            strPart = CStr(varParts(lngPart))
            Select Case True

                ' Check for username/password
                Case StartsWith(strPart, "UID=", vbTextCompare), _
                    StartsWith(strPart, "PWD=", vbTextCompare)
                    ' These values are not needed when using a trusted connection.
                    If (InStr(1, strConnection, "Trusted_Connection=Yes", vbTextCompare) = 0) _
                        Or (Options.SanitizeLevel < eslStandard) Then
                        ' Retain the values if not using trusted connection, or if
                        ' using less than the normal sanitize level.
                        .Add strPart
                    End If

                ' Check database path to convert to relative
                Case StartsWith(strPart, "DATABASE=", vbTextCompare)
                    .Add GetRelativeConnect(strPart)

                ' Add all other sections
                Case Else
                    .Add strPart
            End Select
        Next lngPart

        ' Remove final semicolon added with .AppendOnAdd
        .Remove 1

        ' Return sanitized string
        SanitizeConnectionString = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetRelativeConnect
' Author    : Adam Waller
' Date      : 2/22/2021
' Purpose   : Convert the connection string to a relative path if possible
'---------------------------------------------------------------------------------------
'
Private Function GetRelativeConnect(strConnect As String) As String

    Dim strPath As String
    Dim strAdd As String

    ' Extract the file path from the connection string
    strPath = GetConnectPath(strConnect)

    ' With linked text files, it may just be the folder path. If so, include
    ' a final slash so it can correctly resolve to a relative path.
    If strPath = GetUncPath(CurrentProject.Path) Then strAdd = PathSep

    ' Prefix with the database property name, just in case the same string is used in
    ' another part of the connection string.
    GetRelativeConnect = Replace(strConnect, "DATABASE=" & strPath, _
        "DATABASE=" & GetRelativePath(strPath & strAdd), , , vbTextCompare)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFullConnect
' Author    : Adam Waller
' Date      : 2/22/2021
' Purpose   : Return the full path expanded from a relative one.
'---------------------------------------------------------------------------------------
'
Public Function GetFullConnect(strRelative As String) As String

    Dim strPath As String
    Dim strFull As String

    ' Extract the file path from the connection string
    strPath = GetConnectPath(strRelative)

    ' Convert to full path
    strFull = GetPathFromRelative(strPath)

    ' Remove any trailing slash after conversion from relative.
    If strFull <> strPath And Right$(strFull, 1) = PathSep Then
        strFull = Left$(strFull, Len(strFull) - 1)
    End If

    ' Prefix with the database property name, just in case the same string is used in
    ' another part of the connection string.
    GetFullConnect = Replace(strRelative, "DATABASE=" & strPath, _
        "DATABASE=" & strFull, , , vbTextCompare)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetConnectPath
' Author    : Adam Waller
' Date      : 2/22/2021
' Purpose   : Return embedded connect path
'---------------------------------------------------------------------------------------
'
Private Function GetConnectPath(strConnect As String) As String

    Dim lngPos As Integer
    Dim lngStart As String
    Dim strPath As String

    ' Look for path after "Database" connection property
    lngPos = InStr(1, strConnect, "Database=", vbTextCompare)
    If lngPos > 0 Then
        lngStart = lngPos + Len("Database=")
        lngPos = InStr(lngStart, strConnect, ";")
        If lngPos > 0 Then
            ' Stop at semi-colon
            strPath = Mid$(strConnect, lngStart, lngPos - lngStart)
        Else
            ' Grab the rest of the string
            strPath = Mid$(strConnect, lngStart)
        End If
    End If

    ' Return path, if any
    GetConnectPath = strPath

End Function


'---------------------------------------------------------------------------------------
' Procedure : CacheConnection
' Author    : bclothier
' Date      : 3/31/2023
' Purpose   : Open an ODBC database to allow us to leverage Access' built-in caching
'             and hopefully reduce the numbers of ODBC prompts. Because the connection
'             may be incomplete, we will force a prompt for the user to then fill in
'---------------------------------------------------------------------------------------
'
Public Function CacheConnection(strConnect As String) As Boolean
    If Not (Left$(strConnect, 5) = "ODBC;") Then
        Exit Function
    End If

    Dim qdf As DAO.QueryDef

    If m_dCachedConnections Is Nothing Then
        Set m_dCachedConnections = New Dictionary
    End If

    If m_dCachedConnections.Exists(strConnect) Then
        CacheConnection = True
    Else
        ' We need to use the CurrentDb because it's the one that'll get stuff imported into. Otherwise,
        ' we will get unwanted prompts during the import.
        Set qdf = CurrentDb.CreateQueryDef

        ' There is a bug where if Name property is left uninitialized or has a null string passed into it
        ' the query will not behave correctly and gives strange error. Setting it to an initialized empty
        ' string will avoid the bug.
        qdf.Name = ""

        ' We must provide a SQL statement. Every database engine understand this, right?
        qdf.SQL = "SELECT 1;"
        qdf.Connect = strConnect

        LogUnhandledErrors
        On Error Resume Next
        qdf.OpenRecordset
        If Err.Number Then
            Set qdf = Nothing
        End If
        On Error GoTo 0

        If Not qdf Is Nothing Then
            m_dCachedConnections.Add strConnect, qdf
            CacheConnection = True
        End If
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : CacheConnection
' Author    : bclothier
' Date      : 3/31/2023
' Purpose   : Closes all cached connections
'---------------------------------------------------------------------------------------
'
Public Sub CloseCachedConnections()
    If m_dCachedConnections Is Nothing Then
        Exit Sub
    End If

    If m_dCachedConnections.Count Then
        Dim qdf As DAO.QueryDef
        Dim varKey As Variant

        For Each varKey In m_dCachedConnections.Keys
            Set qdf = m_dCachedConnections.Item(varKey)
            qdf.Close
            Set qdf = Nothing
            m_dCachedConnections.Remove varKey
        Next
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CacheBackEndConnections
' Author    : Adam Waller
' Date      : 03/11/2026
' Purpose   : Open persistent read-only connections to Access back-end database files
'           : linked from the current database. Holding these DAO.Database references
'           : keeps the Jet/ACE connection pool warm, avoiding repeated open/close
'           : cycles when exporting linked table metadata and data.
'---------------------------------------------------------------------------------------
'
Public Sub CacheBackEndConnections()

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim strConnect As String
    Dim strPath As String
    Dim strPwd As String
    Dim dbBackEnd As DAO.Database
    Dim varKey As Variant

    Perf.OperationStart "Cache Back-End Connections"

    If m_dBackEndConnections Is Nothing Then
        Set m_dBackEndConnections = New Dictionary
    End If
    If m_dUnavailableBackEnds Is Nothing Then
        Set m_dUnavailableBackEnds = New Dictionary
    End If

    ' Scan all table definitions for Access back-end links
    Set dbs = CurrentDb
    For Each tdf In dbs.TableDefs

        ' Only process Access back-end linked tables (;DATABASE=... or MS Access;...)
        strConnect = tdf.Connect
        If Len(strConnect) > 0 Then
            If InStr(1, strConnect, ";DATABASE=", vbTextCompare) = 1 _
                Or InStr(1, strConnect, "MS Access;", vbTextCompare) = 1 Then

                ' Extract the back-end file path
                strPath = GetConnectPath(strConnect)
                If Len(strPath) > 0 Then
                    If m_dBackEndConnections.Exists(strPath) Then
                        ' Already cached successfully
                    ElseIf m_dUnavailableBackEnds.Exists(UCase$(strPath)) Then
                        ' Already known unavailable - increment table count
                        m_dUnavailableBackEnds(UCase$(strPath)) = _
                            m_dUnavailableBackEnds(UCase$(strPath)) + 1
                    Else
                        ' Open the back-end database in shared, read-only mode
                        strPwd = GetConnectPart(strConnect, "PWD")
                        LogUnhandledErrors
                        On Error Resume Next
                        If Len(strPwd) > 0 Then
                            Set dbBackEnd = DBEngine.OpenDatabase(strPath, False, True, ";PWD=" & strPwd)
                        Else
                            Set dbBackEnd = DBEngine.OpenDatabase(strPath, False, True)
                        End If
                        If Err.Number = 0 Then
                            m_dBackEndConnections.Add strPath, dbBackEnd
                        Else
                            Err.Clear
                            m_dUnavailableBackEnds.Add UCase$(strPath), 1
                        End If
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
    Next tdf

    ' Log the cached connections
    If m_dBackEndConnections.Count > 0 Then
        Log.Add T("Caching {0} back-end database connection(s)", _
            var0:=m_dBackEndConnections.Count), Options.ShowDebug
        For Each varKey In m_dBackEndConnections.Keys
            Log.Add "  " & CStr(varKey), Options.ShowDebug
        Next varKey
    End If

    ' Log a single warning per unavailable back-end
    For Each varKey In m_dUnavailableBackEnds.Keys
        Log.Error eelWarning, T("Back-end database unavailable: {0} ({1} linked table(s) will be skipped)", _
            var0:=CStr(varKey), var1:=m_dUnavailableBackEnds(varKey)), _
            ModuleName & ".CacheBackEndConnections"
    Next varKey

    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CloseBackEndConnections
' Author    : Adam Waller
' Date      : 03/11/2026
' Purpose   : Close all cached back-end database connections opened by
'           : CacheBackEndConnections.
'---------------------------------------------------------------------------------------
'
Public Sub CloseBackEndConnections()

    Dim varKey As Variant
    Dim dbBackEnd As DAO.Database

    If m_dBackEndConnections Is Nothing Then Exit Sub
    If m_dBackEndConnections.Count = 0 Then Exit Sub

    LogUnhandledErrors
    On Error Resume Next

    For Each varKey In m_dBackEndConnections.Keys
        Set dbBackEnd = m_dBackEndConnections.Item(varKey)
        dbBackEnd.Close
        Set dbBackEnd = Nothing
    Next varKey

    On Error GoTo 0

    Set m_dBackEndConnections = Nothing
    Set m_dUnavailableBackEnds = Nothing

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsBackEndUnavailable
' Author    : Adam Waller
' Date      : 03/11/2026
' Purpose   : Returns True if the back-end identified by the given connection string
'           : has been marked as unavailable. Used to skip linked tables whose back-end
'           : database or server is known to be unreachable.
'---------------------------------------------------------------------------------------
'
Public Function IsBackEndUnavailable(strConnect As String) As Boolean

    Dim strKey As String

    If m_dUnavailableBackEnds Is Nothing Then Exit Function
    If m_dUnavailableBackEnds.Count = 0 Then Exit Function

    strKey = GetBackEndKey(strConnect)
    If Len(strKey) > 0 Then
        IsBackEndUnavailable = m_dUnavailableBackEnds.Exists(strKey)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : MarkBackEndUnavailable
' Author    : Adam Waller
' Date      : 03/11/2026
' Purpose   : Record a back-end as unavailable and log a single warning. Called
'           : reactively when a linked table export fails and the back-end connection
'           : test also fails, indicating the entire back-end is down rather than a
'           : single table being missing.
'---------------------------------------------------------------------------------------
'
Public Sub MarkBackEndUnavailable(strConnect As String, strTableName As String)

    Dim strKey As String

    If m_dUnavailableBackEnds Is Nothing Then
        Set m_dUnavailableBackEnds = New Dictionary
    End If

    strKey = GetBackEndKey(strConnect)
    If Len(strKey) = 0 Then Exit Sub

    If Not m_dUnavailableBackEnds.Exists(strKey) Then
        m_dUnavailableBackEnds.Add strKey, 1
        Log.Error eelWarning, T("Back-end connection unavailable: {0}. " & _
            "Linked tables from this source will be skipped.", _
            var0:=strKey), ModuleName & ".MarkBackEndUnavailable"
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : TestBackEndConnection
' Author    : Adam Waller
' Date      : 03/11/2026
' Purpose   : Perform a lightweight server-level connection test for non-Access
'           : connections (ODBC, etc.) by creating a temp QueryDef with SELECT 1.
'           : Returns True if the server responds. Used to distinguish "server down"
'           : from "single table missing" when a linked table fails TableExists.
'---------------------------------------------------------------------------------------
'
Public Function TestBackEndConnection(strConnect As String) As Boolean

    Dim qdf As DAO.QueryDef

    ' Only test non-Access connections; Access back-ends are proactively
    ' tested in CacheBackEndConnections via DBEngine.OpenDatabase.
    If InStr(1, strConnect, ";DATABASE=", vbTextCompare) = 1 _
        Or InStr(1, strConnect, "MS Access;", vbTextCompare) = 1 Then
        ' Access back-end: check against the cached connections dictionary
        Dim strPath As String
        strPath = GetConnectPath(strConnect)
        If Len(strPath) > 0 Then
            If Not m_dBackEndConnections Is Nothing Then
                TestBackEndConnection = m_dBackEndConnections.Exists(strPath)
            End If
        End If
        Exit Function
    End If

    ' For ODBC and other connection types, attempt a lightweight query
    LogUnhandledErrors
    On Error Resume Next

    Set qdf = CurrentDb.CreateQueryDef("")
    qdf.Connect = strConnect
    qdf.SQL = "SELECT 1;"
    qdf.OpenRecordset
    TestBackEndConnection = (Err.Number = 0)

    If Not qdf Is Nothing Then
        qdf.Close
        Set qdf = Nothing
    End If

    On Error GoTo 0

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetBackEndKey
' Author    : Adam Waller
' Date      : 03/11/2026
' Purpose   : Normalize a connection string to a back-end identifier suitable for
'           : dictionary key comparison. Returns file path for Access back-ends,
'           : DSN or DRIVER+SERVER+DATABASE for ODBC, or the full string for others.
'---------------------------------------------------------------------------------------
'
Private Function GetBackEndKey(strConnect As String) As String

    Dim strDSN As String

    If Len(strConnect) = 0 Then Exit Function

    If InStr(1, strConnect, ";DATABASE=", vbTextCompare) = 1 _
        Or InStr(1, strConnect, "MS Access;", vbTextCompare) = 1 Then
        ' Access back-end: use the file path as the key
        GetBackEndKey = UCase$(GetConnectPath(strConnect))

    ElseIf InStr(1, strConnect, "ODBC;", vbTextCompare) = 1 Then
        ' ODBC connection: build key from DSN or DRIVER+SERVER+DATABASE
        strDSN = GetConnectPart(strConnect, "DSN")
        If Len(strDSN) > 0 Then
            GetBackEndKey = "ODBC:DSN=" & UCase$(strDSN)
        Else
            GetBackEndKey = "ODBC:" & _
                UCase$(GetConnectPart(strConnect, "DRIVER")) & ";" & _
                UCase$(GetConnectPart(strConnect, "SERVER")) & ";" & _
                UCase$(GetConnectPart(strConnect, "DATABASE"))
        End If

    Else
        ' Other connection types: use full string normalized to uppercase
        GetBackEndKey = UCase$(strConnect)
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetConnectPart
' Author    : Adam Waller
' Date      : 03/11/2026
' Purpose   : Extract a named part from a semicolon-delimited connection string.
'           : E.g. GetConnectPart("ODBC;DRIVER={SQL Server};SERVER=mysvr", "SERVER")
'           : returns "mysvr".
'---------------------------------------------------------------------------------------
'
Public Function GetConnectPart(strConnect As String, strPart As String) As String

    Dim lngStart As Long
    Dim lngEnd As Long

    lngStart = InStr(1, strConnect, strPart & "=", vbTextCompare)
    If lngStart > 0 Then
        lngStart = lngStart + Len(strPart) + 1
        lngEnd = InStr(lngStart, strConnect, ";")
        If lngEnd > 0 Then
            GetConnectPart = Mid$(strConnect, lngStart, lngEnd - lngStart)
        Else
            GetConnectPart = Mid$(strConnect, lngStart)
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetConnectionEnvKey
' Author    : Adam Waller
' Date      : 03/17/2026
' Purpose   : Generate a deterministic, readable .env key for a connection string
'           : based on the database identity (DATABASE= value or DSN= value).
'           : Volatile parts (SERVER=, DRIVER=, credentials) are excluded so the
'           : key is stable across environments. Auto-generated keys are lowercase.
'           : Falls back to a hash when no identity can be extracted.
'---------------------------------------------------------------------------------------
'
Public Function GetConnectionEnvKey(strConnect As String) As String

    Dim strIdentity As String
    Dim strDbName As String
    Dim varParts As Variant
    Dim lngPart As Long
    Dim strPart As String

    If Len(strConnect) = 0 Then Exit Function

    ' Extract database identity based on connection type
    strDbName = GetConnectPart(strConnect, "DATABASE")
    If Len(strDbName) > 0 Then
        If InStr(1, strConnect, ";DATABASE=", vbTextCompare) = 1 _
            Or InStr(1, strConnect, "MS Access;", vbTextCompare) = 1 Then
            ' Access back-end: use filename without extension
            strIdentity = FSO.GetBaseName(strDbName)
        Else
            ' ODBC or other: use DATABASE name directly
            strIdentity = strDbName
        End If
    Else
        ' No DATABASE=; try DSN=
        strIdentity = GetConnectPart(strConnect, "DSN")
    End If

    If Len(strIdentity) > 0 Then
        ' Readable key from database identity (sanitized to lowercase)
        GetConnectionEnvKey = ENV_KEY_PREFIX & SanitizeKeyName(strIdentity)
    Else
        ' Fallback: hash the connection string excluding credentials
        varParts = Split(strConnect, ";")
        With New clsConcat
            .AppendOnAdd = ";"
            For lngPart = 0 To UBound(varParts)
                strPart = CStr(varParts(lngPart))
                If Not StartsWith(strPart, "UID=", vbTextCompare) _
                    And Not StartsWith(strPart, "PWD=", vbTextCompare) Then
                    .Add strPart
                End If
            Next lngPart
            .Remove 1
            GetConnectionEnvKey = ENV_KEY_PREFIX & GetSimpleHash(.GetStr)
        End With
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : SanitizeKeyName
' Author    : Adam Waller
' Date      : 04/09/2026
' Purpose   : Replace non-alphanumeric characters (except underscore) with underscores
'           : to produce a valid .env key name.
'---------------------------------------------------------------------------------------
'
Private Function SanitizeKeyName(strName As String) As String

    Dim lngPos As Long
    Dim strChar As String

    SanitizeKeyName = LCase$(strName)
    For lngPos = 1 To Len(SanitizeKeyName)
        strChar = Mid$(SanitizeKeyName, lngPos, 1)
        Select Case True
            Case (strChar >= "a" And strChar <= "z"), _
                 (strChar >= "0" And strChar <= "9"), _
                 (strChar = "_")
                ' Keep lowercase alphanumeric and underscore
            Case Else
                Mid$(SanitizeKeyName, lngPos, 1) = "_"
        End Select
    Next lngPos

End Function


'---------------------------------------------------------------------------------------
' Procedure : ShouldUseEnvForConnection
' Author    : Adam Waller
' Date      : 03/17/2026
' Purpose   : Determines whether a connection string should be stored in the .env file.
'           : First checks for user-defined named connections (Tier 2), then falls back
'           : to auto-generated keys (Tier 1) based on the UseEnvForConnections option.
'           : Returns the env key if substitution should happen, or empty string if not.
'---------------------------------------------------------------------------------------
'
Public Function ShouldUseEnvForConnection(strConnect As String) As String

    Dim strNamedKey As String

    ' Must be on export format v5+
    If Options.ExportFormatVersion < EFV_5_0_0 Then Exit Function

    ' Tier 2: Check user-defined named connections first
    strNamedKey = FindNamedConnectionKey(strConnect)
    If Len(strNamedKey) > 0 Then
        ShouldUseEnvForConnection = strNamedKey
        Exit Function
    End If

    ' Tier 1: Auto-generated keys
    Select Case Options.UseEnvForConnections
        Case uecNever
            Exit Function

        Case uecAlways
            If Len(strConnect) > 0 Then
                ShouldUseEnvForConnection = GetConnectionEnvKey(strConnect)
            End If

        Case uecAuto
            ' Only use .env when credentials are detected
            If HasCredentials(strConnect) Then
                ShouldUseEnvForConnection = GetConnectionEnvKey(strConnect)
            End If

    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : HasCredentials
' Author    : Adam Waller
' Date      : 03/17/2026
' Purpose   : Returns True if the connection string contains UID= or PWD= patterns.
'---------------------------------------------------------------------------------------
'
Private Function HasCredentials(strConnect As String) As Boolean
    HasCredentials = (InStr(1, strConnect, "PWD=", vbTextCompare) > 0) _
        Or (InStr(1, strConnect, "UID=", vbTextCompare) > 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveConnectionToEnv
' Author    : Adam Waller
' Date      : 03/17/2026
' Purpose   : Save a connection string to the root .env file using the specified key.
'           : Adds a comment with the source object name when creating a new entry.
'           : Tracks written keys for unused entry detection.
'---------------------------------------------------------------------------------------
'
Public Sub SaveConnectionToEnv(strKey As String, strConnect As String, strSourceObject As String)

    Dim cEnv As clsDotEnv
    Dim strFile As String
    Dim strCommentKey As String
    Dim blnNewFile As Boolean

    ' Track this key as written during the current export
    If m_dEnvKeysWritten Is Nothing Then Set m_dEnvKeysWritten = New Dictionary
    If Not m_dEnvKeysWritten.Exists(strKey) Then
        m_dEnvKeysWritten.Add strKey, strSourceObject
    End If

    ' For named connections (Tier 2), don't overwrite the user's .env value.
    ' The key is tracked above for unused-entry detection.
    If IsDefinedConnectionName(strKey) Then Exit Sub

    strFile = GetEnvFilePath
    Set cEnv = GetEnvCache(True)

    ' Check if this is a brand new file (no existing entries)
    blnNewFile = (cEnv.Lines.Count = 0)

    If blnNewFile Then
        ' Add header comments for new files
        cEnv.Lines.Add "COMMENT_H1", "# Connection strings for linked tables and pass-through queries."
        cEnv.Lines.Add "COMMENT_H2", "# This file should be excluded from version control (.gitignore)."
        cEnv.Lines.Add "COMMENT_H3", "# Each developer maintains their own copy with their credentials."
        cEnv.Lines.Add "COMMENT_H4", ""
    End If

    ' Add comment and value for new entries
    If Not cEnv.Lines.Exists(strKey) Then
        strCommentKey = "COMMENT_" & Format(cEnv.Lines.Count, "###")
        cEnv.Lines.Add strCommentKey, "# " & strSourceObject
    End If

    ' Set or update the value
    cEnv.SetVar strKey, strConnect

    ' Save the file
    cEnv.SaveToFile strFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResolveEnvConnection
' Author    : Adam Waller
' Date      : 03/17/2026
' Purpose   : Given a string like "env:conn_appdb", look up the key in the .env file
'           : and return the full connection string. Returns empty string with a warning
'           : if the key is not found.
'---------------------------------------------------------------------------------------
'
Public Function ResolveEnvConnection(strRef As String) As String

    Dim strKey As String
    Dim strValue As String
    Dim cEnv As clsDotEnv

    If Not IsEnvReference(strRef) Then Exit Function

    ' Extract the key from the "env:" prefix
    strKey = Mid$(strRef, 5)

    Set cEnv = GetEnvCache(False)
    strValue = cEnv.GetVar(strKey, False)

    If Len(strValue) > 0 Then
        ResolveEnvConnection = strValue
    Else
        ' Track missing keys and only warn on first occurrence
        If m_dMissingEnvKeys Is Nothing Then Set m_dMissingEnvKeys = New Dictionary
        If Not m_dMissingEnvKeys.Exists(strKey) Then
            m_dMissingEnvKeys.Add strKey, 1
            Log.Error eelWarning, T("Connection key not found in .env file: {0}", _
                var0:=strKey), ModuleName & ".ResolveEnvConnection"
        Else
            m_dMissingEnvKeys(strKey) = m_dMissingEnvKeys(strKey) + 1
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsEnvReference
' Author    : Adam Waller
' Date      : 03/17/2026
' Purpose   : Returns True if the string is an env: reference to a .env file entry.
'---------------------------------------------------------------------------------------
'
Public Function IsEnvReference(strConnect As String) As Boolean
    IsEnvReference = StartsWith(strConnect, "env:", vbTextCompare)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ResolveEnvReferencesInText
' Author    : Adam Waller
' Date      : 03/17/2026
' Purpose   : Find and replace all env:conn_* patterns in text content.
'           : Used to resolve .env references in .qdef files before LoadFromText.
'---------------------------------------------------------------------------------------
'
Public Function ResolveEnvReferencesInText(strContent As String) As String

    Dim lngPos As Long
    Dim lngEnd As Long
    Dim strRef As String
    Dim strResolved As String

    ResolveEnvReferencesInText = strContent

    ' Find all occurrences of "env:conn_" and resolve them
    lngPos = InStr(1, ResolveEnvReferencesInText, "env:" & ENV_KEY_PREFIX, vbTextCompare)
    Do While lngPos > 0
        ' Find the end of the reference (next quote or semicolon or end of line)
        lngEnd = lngPos + 5 + Len(ENV_KEY_PREFIX)
        Do While lngEnd <= Len(ResolveEnvReferencesInText)
            Select Case Mid$(ResolveEnvReferencesInText, lngEnd, 1)
                Case """", ";", vbCr, vbLf
                    Exit Do
                Case Else
                    lngEnd = lngEnd + 1
            End Select
        Loop

        ' Extract and resolve the reference
        strRef = Mid$(ResolveEnvReferencesInText, lngPos, lngEnd - lngPos)
        strResolved = ResolveEnvConnection(strRef)

        If Len(strResolved) > 0 Then
            ResolveEnvReferencesInText = Left$(ResolveEnvReferencesInText, lngPos - 1) & _
                strResolved & Mid$(ResolveEnvReferencesInText, lngEnd)
            ' Advance past the resolved string
            lngPos = lngPos + Len(strResolved)
        Else
            ' Could not resolve; leave as-is and move past it
            lngPos = lngEnd
        End If

        ' Find next occurrence
        lngPos = InStr(lngPos, ResolveEnvReferencesInText, "env:" & ENV_KEY_PREFIX, vbTextCompare)
    Loop

End Function


'---------------------------------------------------------------------------------------
' Procedure : FindNamedConnectionKey
' Author    : Adam Waller
' Date      : 04/09/2026
' Purpose   : Check if a connection string matches any user-defined named connection
'           : from EnvConnectionNames in vcs-options.json. Compares the raw connection
'           : against the expanded .env values using order-independent parameter matching.
'           : Returns the named key if matched, or empty string if not.
'---------------------------------------------------------------------------------------
'
Public Function FindNamedConnectionKey(strConnect As String) As String

    Dim colNames As Collection
    Dim varName As Variant
    Dim strName As String
    Dim strEnvValue As String
    Dim cEnv As clsDotEnv

    Set colNames = Options.EnvConnectionNames
    If colNames Is Nothing Then Exit Function
    If colNames.Count = 0 Then Exit Function

    Set cEnv = GetEnvCache(True)

    For Each varName In colNames
        strName = CStr(varName)
        strEnvValue = cEnv.GetVar(strName, False)
        If Len(strEnvValue) > 0 Then
            ' Expand relative paths so both sides are comparable
            If ConnectionParamsMatch(strConnect, GetFullConnect(strEnvValue)) Then
                FindNamedConnectionKey = strName
                Exit Function
            End If
        End If
    Next varName

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsDefinedConnectionName
' Author    : Adam Waller
' Date      : 04/09/2026
' Purpose   : Returns True if the key is one of the user-defined named connections
'           : from EnvConnectionNames in vcs-options.json.
'---------------------------------------------------------------------------------------
'
Private Function IsDefinedConnectionName(strKey As String) As Boolean

    Dim colNames As Collection
    Dim varName As Variant

    Set colNames = Options.EnvConnectionNames
    If colNames Is Nothing Then Exit Function

    For Each varName In colNames
        If StrComp(CStr(varName), strKey, vbTextCompare) = 0 Then
            IsDefinedConnectionName = True
            Exit Function
        End If
    Next varName

End Function


'---------------------------------------------------------------------------------------
' Procedure : ConnectionParamsMatch
' Author    : Adam Waller
' Date      : 04/09/2026
' Purpose   : Compare two connection strings by their parameter values, ignoring
'           : parameter order. Both keys and values are compared case-insensitively.
'---------------------------------------------------------------------------------------
'
Private Function ConnectionParamsMatch(strConnect1 As String, strConnect2 As String) As Boolean

    Dim d1 As Dictionary
    Dim d2 As Dictionary
    Dim varKey As Variant

    Set d1 = ParseConnectionParams(strConnect1)
    Set d2 = ParseConnectionParams(strConnect2)

    If d1.Count <> d2.Count Then Exit Function

    For Each varKey In d1.Keys
        If Not d2.Exists(CStr(varKey)) Then Exit Function
        If StrComp(CStr(d1(varKey)), CStr(d2(varKey)), vbTextCompare) <> 0 Then Exit Function
    Next varKey

    ConnectionParamsMatch = True

End Function


'---------------------------------------------------------------------------------------
' Procedure : ParseConnectionParams
' Author    : Adam Waller
' Date      : 04/09/2026
' Purpose   : Parse a semicolon-delimited connection string into a dictionary of
'           : parameter name/value pairs. Prefix segments without "=" (like "ODBC")
'           : are stored as keys with empty values.
'---------------------------------------------------------------------------------------
'
Private Function ParseConnectionParams(strConnect As String) As Dictionary

    Dim dParams As Dictionary
    Dim varParts As Variant
    Dim lngPart As Long
    Dim strPart As String
    Dim lngEq As Long

    Set dParams = New Dictionary
    dParams.CompareMode = TextCompare

    If Len(strConnect) = 0 Then
        Set ParseConnectionParams = dParams
        Exit Function
    End If

    varParts = Split(strConnect, ";")
    For lngPart = 0 To UBound(varParts)
        strPart = Trim$(CStr(varParts(lngPart)))
        If Len(strPart) > 0 Then
            lngEq = InStr(1, strPart, "=")
            If lngEq > 0 Then
                dParams(Left$(strPart, lngEq - 1)) = Mid$(strPart, lngEq + 1)
            Else
                dParams(strPart) = vbNullString
            End If
        End If
    Next lngPart

    Set ParseConnectionParams = dParams

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetEnvFilePath
' Author    : Adam Waller
' Date      : 03/17/2026
' Purpose   : Return the full path to the root .env file in the export folder.
'---------------------------------------------------------------------------------------
'
Public Function GetEnvFilePath() As String
    GetEnvFilePath = Options.GetExportFolder & ".env"
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetEnvCache
' Author    : Adam Waller
' Date      : 03/17/2026
' Purpose   : Return a cached clsDotEnv instance to avoid re-reading the .env file
'           : for every table/query during an export or import operation.
'           : Set blnForWrite to True during export to allow creating a new file.
'---------------------------------------------------------------------------------------
'
Private Function GetEnvCache(blnForWrite As Boolean) As clsDotEnv

    Dim strFile As String

    If m_cEnvCache Is Nothing Then
        Set m_cEnvCache = New clsDotEnv
        strFile = GetEnvFilePath
        If FSO.FileExists(strFile) Then
            m_cEnvCache.LoadFromFile strFile
        End If
    End If

    Set GetEnvCache = m_cEnvCache

End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearEnvCache
' Author    : Adam Waller
' Date      : 03/17/2026
' Purpose   : Clear the cached .env instance after an export or import is complete.
'---------------------------------------------------------------------------------------
'
Public Sub ClearEnvCache()
    Set m_cEnvCache = Nothing
    Set m_dEnvKeysWritten = Nothing
    Set m_dMissingEnvKeys = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetMissingEnvKeys
' Author    : Adam Waller
' Date      : 04/09/2026
' Purpose   : Return dictionary of .env keys that were referenced but not found.
'           : Keys are the key names, values are the number of times each was requested.
'---------------------------------------------------------------------------------------
'
Public Function GetMissingEnvKeys() As Dictionary
    If m_dMissingEnvKeys Is Nothing Then
        Set GetMissingEnvKeys = New Dictionary
    Else
        Set GetMissingEnvKeys = m_dMissingEnvKeys
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : LogUnusedEnvEntries
' Author    : Adam Waller
' Date      : 03/17/2026
' Purpose   : After a full export, compare written keys against existing .env entries
'           : and log any unused conn_* key names to the log file (not console).
'---------------------------------------------------------------------------------------
'
Public Sub LogUnusedEnvEntries()

    Dim cEnv As clsDotEnv
    Dim varKey As Variant
    Dim strKey As String
    Dim lngCount As Long

    If m_dEnvKeysWritten Is Nothing Then Exit Sub

    Set cEnv = GetEnvCache(False)

    For Each varKey In cEnv.Lines.Keys
        strKey = CStr(varKey)
        If StartsWith(strKey, ENV_KEY_PREFIX, vbTextCompare) Then
            If Not m_dEnvKeysWritten.Exists(strKey) Then
                If lngCount = 0 Then
                    Log.Add T("Unused connection entries in .env file:"), False
                End If
                Log.Add "  " & strKey, False
                lngCount = lngCount + 1
            End If
        End If
    Next varKey

End Sub


