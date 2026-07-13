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
Private Const APP_ENV_KEY As String = "APP_ENV"
Private m_dCachedConnections As Dictionary
Private m_dBackEndConnections As Dictionary
Private m_dUnavailableBackEnds As Dictionary
Private m_cEnvCache As clsDotEnv
Private m_cEnvResolved As clsDotEnv
Private m_dEnvKeysWritten As Dictionary
Private m_dMissingEnvKeys As Dictionary
Private m_dConnState As Dictionary
Private m_dStrippedConnWarn As Dictionary


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
' Procedure : ClearConnState
' Author    : Adam Waller
' Date      : 06/19/2026
' Purpose   : Reset per-connection build/import state at the start of a build.
'---------------------------------------------------------------------------------------
'
Public Sub ClearConnState()
    Set m_dConnState = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetConnState
' Author    : Adam Waller
' Date      : 06/19/2026
' Purpose   : Return (and create if needed) the state dictionary for a connection key.
'---------------------------------------------------------------------------------------
'
Public Function GetConnState(strKey As String) As Dictionary

    Dim dState As Dictionary

    If Len(strKey) = 0 Then Exit Function

    If m_dConnState Is Nothing Then Set m_dConnState = New Dictionary

    If m_dConnState.Exists(strKey) Then
        Set GetConnState = m_dConnState(strKey)
    Else
        Set dState = New Dictionary
        dState.Add "Status", csUnknown
        dState.Add "Completed", vbNullString
        dState.Add "EnvKey", vbNullString
        dState.Add "Source", vbNullString
        dState.Add "NeedsSave", False
        m_dConnState.Add strKey, dState
        Set GetConnState = dState
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsConnectionError
' Author    : Adam Waller
' Date      : 06/19/2026
' Purpose   : Returns True for ODBC/login errors that warrant Retry/Ignore/Abort.
'---------------------------------------------------------------------------------------
'
Public Function IsConnectionError(lngErr As Long) As Boolean
    Select Case lngErr
        Case 3059, 3146, 3151
            IsConnectionError = True
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsOracleOdbcConnect
' Author    : Adam Waller
' Date      : 07/13/2026
' Purpose   : Returns True when an ODBC connection string names an Oracle driver.
'           : DSN-only strings without DRIVER= are not detected (see DECISIONS.md).
'---------------------------------------------------------------------------------------
'
Public Function IsOracleOdbcConnect(strConnect As String) As Boolean
    IsOracleOdbcConnect = (InStr(1, GetConnectPart(strConnect, "DRIVER"), "Oracle", vbTextCompare) > 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetConnectivityProbeSql
' Author    : Adam Waller
' Date      : 07/13/2026
' Purpose   : Lightweight SQL for ODBC connectivity probes. Oracle requires FROM DUAL.
'---------------------------------------------------------------------------------------
'
Public Function GetConnectivityProbeSql(strConnect As String) As String
    If IsOracleOdbcConnect(strConnect) Then
        GetConnectivityProbeSql = "SELECT 1 FROM DUAL;"
    Else
        GetConnectivityProbeSql = "SELECT 1;"
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : HandleConnectionFailure
' Author    : Adam Waller
' Date      : 06/19/2026
' Purpose   : Prompt the user to Retry, Ignore, or Abort after a connection failure.
'---------------------------------------------------------------------------------------
'
Public Function HandleConnectionFailure(strKey As String, strObjectLabel As String, _
    lngErr As Long, strErrDesc As String) As eConnFailAction

    Dim strDetail As String
    Dim intResult As VbMsgBoxResult

    If Len(strErrDesc) > 0 Then
        strDetail = vbCrLf & vbCrLf & lngErr & " - " & strErrDesc
    End If

    intResult = MsgBox2(T("Connection Failed"), _
        T("Could not connect for {0}.", var0:=strObjectLabel) & strDetail, _
        T("Retry to try again, Ignore to skip objects on this connection, or Abort to cancel the build."), _
        vbAbortRetryIgnore Or vbExclamation, , vbAbort)

    Select Case intResult
        Case vbRetry
            HandleConnectionFailure = cfaRetry
        Case vbIgnore
            GetConnState(strKey)("Status") = csIgnored
            HandleConnectionFailure = cfaIgnore
        Case Else
            HandleConnectionFailure = cfaAbort
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : RecordAuthenticatedConnection
' Author    : Adam Waller
' Date      : 06/19/2026
' Purpose   : Remember a successfully authenticated connection for reuse and optional
'           : .env save at the end of the build.
'---------------------------------------------------------------------------------------
'
Public Sub RecordAuthenticatedConnection(strKey As String, strEnvKey As String, _
    strCompletedConnect As String, strSource As String)

    Dim dState As Dictionary
    Dim strEnvValue As String
    Dim cEnv As clsDotEnv

    If Len(strKey) = 0 Then Exit Sub
    If Len(strCompletedConnect) = 0 Then Exit Sub

    Set dState = GetConnState(strKey)
    dState("Status") = csOK
    dState("Completed") = strCompletedConnect

    If Len(strEnvKey) = 0 Then Exit Sub
    If Options.UseEnvForConnections = uecNever Then Exit Sub

    ' Only worth saving if the completed string carries the authentication detail
    ' (UID/PWD, trusted, or an AD/SSPI authentication method) that lets a future
    ' build connect without prompting.
    If Not HasAuthInfo(strCompletedConnect) Then Exit Sub

    Set cEnv = GetEnvCache(False)
    strEnvValue = cEnv.GetVar(strEnvKey, False)
    If Len(strEnvValue) > 0 Then
        strEnvValue = GetFullConnect(strEnvValue)
        ' If the stored value already resolves authentication on its own, it would
        ' not prompt, so there is nothing new to save.
        If HasAuthInfo(strEnvValue) Then Exit Sub
        ' Identical parameters mean nothing meaningful changed.
        If ConnectionParamsMatch(strCompletedConnect, strEnvValue) Then Exit Sub
    End If

    dState("NeedsSave") = True
    dState("EnvKey") = strEnvKey
    dState("Source") = strSource

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PromptAndSaveConnections
' Author    : Adam Waller
' Date      : 06/19/2026
' Purpose   : At the end of a build, save newly authenticated connections to .env per
'           : the UseEnvForConnections policy (Never=off, Auto=prompt, Always=silent).
'---------------------------------------------------------------------------------------
'
Public Sub PromptAndSaveConnections()

    Dim varKey As Variant
    Dim dState As Dictionary
    Dim colSave As New Collection
    Dim strList As String
    Dim lngCount As Long

    If m_dConnState Is Nothing Then Exit Sub

    For Each varKey In m_dConnState.Keys
        Set dState = m_dConnState(varKey)
        If dState("NeedsSave") Then colSave.Add CStr(varKey)
    Next varKey

    If colSave.Count = 0 Then
        Set m_dConnState = Nothing
        Exit Sub
    End If

    If Options.UseEnvForConnections = uecNever Then
        Set m_dConnState = Nothing
        Exit Sub
    End If

    With New clsConcat
        .AppendOnAdd = vbCrLf
        For Each varKey In colSave
            Set dState = m_dConnState(CStr(varKey))
            .Add CStr(dState("EnvKey")) & " (" & CStr(dState("Source")) & ")"
        Next varKey
        strList = .GetStr
    End With

    If Options.UseEnvForConnections = uecAuto Then
        If MsgBox2(T("Save Connection Credentials?"), strList, _
            T("Save them to your .env file so you are not prompted again?"), _
            vbYesNo Or vbQuestion, , vbNo) <> vbYes Then
            Set m_dConnState = Nothing
            Exit Sub
        End If
    End If

    For Each varKey In colSave
        Set dState = m_dConnState(CStr(varKey))
        SaveConnectionToEnv CStr(dState("EnvKey")), CStr(dState("Completed")), CStr(dState("Source"))
        lngCount = lngCount + 1
    Next varKey

    ClearEnvCache
    If lngCount > 0 Then
        Log.Add T("Saved {0} connection credential(s) to .env.", var0:=lngCount), False
    End If
    Set m_dConnState = Nothing

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CacheConnection
' Author    : bclothier
' Date      : 3/31/2023
' Purpose   : Open an ODBC database to allow us to leverage Access' built-in caching
'             and hopefully reduce the numbers of ODBC prompts. Because the connection
'             may be incomplete, we will force a prompt for the user to then fill in
'---------------------------------------------------------------------------------------
'
Public Function CacheConnection(strConnect As String, _
    ByRef strCompletedConnect As String, ByRef lngErr As Long, _
    ByRef strErrDesc As String) As Boolean

    Dim qdf As DAO.QueryDef

    lngErr = 0
    strCompletedConnect = vbNullString
    strErrDesc = vbNullString

    If Not (Left$(strConnect, 5) = "ODBC;") Then
        Exit Function
    End If

    If m_dCachedConnections Is Nothing Then
        Set m_dCachedConnections = New Dictionary
    End If

    If m_dCachedConnections.Exists(strConnect) Then
        Set qdf = m_dCachedConnections(strConnect)
        strCompletedConnect = qdf.Connect
        CacheConnection = True
    Else
        ' We need to use the CurrentDb because it's the one that'll get stuff imported into. Otherwise,
        ' we will get unwanted prompts during the import.
        Set qdf = CurrentDb.CreateQueryDef

        ' There is a bug where if Name property is left uninitialized or has a null string passed into it
        ' the query will not behave correctly and gives strange error. Setting it to an initialized empty
        ' string will avoid the bug.
        qdf.Name = ""

        qdf.SQL = GetConnectivityProbeSql(strConnect)
        qdf.Connect = strConnect

        LogUnhandledErrors
        On Error Resume Next
        qdf.OpenRecordset
        lngErr = Err.Number
        strErrDesc = Err.Description
        On Error GoTo 0

        If lngErr Then
            Set qdf = Nothing
        Else
            strCompletedConnect = qdf.Connect
        End If

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
' Procedure : IsAccessBackEndConnect
' Author    : Adam Waller
' Date      : 06/09/2026
' Purpose   : Returns True if the connection string links to an Access database file.
'---------------------------------------------------------------------------------------
'
Private Function IsAccessBackEndConnect(strConnect As String) As Boolean
    IsAccessBackEndConnect = (InStr(1, strConnect, ";DATABASE=", vbTextCompare) = 1 _
        Or InStr(1, strConnect, "MS Access;", vbTextCompare) = 1)
End Function


'---------------------------------------------------------------------------------------
' Procedure : OpenAccessBackEnd
' Author    : Adam Waller
' Date      : 06/09/2026
' Purpose   : Open an Access back-end database file in shared read-only mode. Returns
'           : Nothing when the file cannot be opened.
'---------------------------------------------------------------------------------------
'
Private Function OpenAccessBackEnd(strConnect As String) As DAO.Database

    Dim strPath As String
    Dim strPwd As String
    Dim dbBackEnd As DAO.Database

    strPath = GetConnectPath(strConnect)
    If Len(strPath) = 0 Then Exit Function

    strPwd = GetConnectPart(strConnect, "PWD")
    LogUnhandledErrors
    On Error Resume Next
    If Len(strPwd) > 0 Then
        Set dbBackEnd = DBEngine.OpenDatabase(strPath, False, True, ";PWD=" & strPwd)
    Else
        Set dbBackEnd = DBEngine.OpenDatabase(strPath, False, True)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        Set dbBackEnd = Nothing
    End If
    On Error GoTo 0

    Set OpenAccessBackEnd = dbBackEnd

End Function


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
    Dim strKey As String
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

        strConnect = tdf.Connect
        If Len(strConnect) > 0 Then
            If IsAccessBackEndConnect(strConnect) Then

                strPath = GetConnectPath(strConnect)
                If Len(strPath) > 0 Then
                    strKey = UCase$(strPath)
                    If m_dBackEndConnections.Exists(strKey) Then
                        ' Already cached successfully
                    ElseIf m_dUnavailableBackEnds.Exists(strKey) Then
                        m_dUnavailableBackEnds(strKey) = m_dUnavailableBackEnds(strKey) + 1
                    Else
                        Set dbBackEnd = OpenAccessBackEnd(strConnect)
                        If Not dbBackEnd Is Nothing Then
                            m_dBackEndConnections.Add strKey, dbBackEnd
                        Else
                            m_dUnavailableBackEnds.Add strKey, 1
                        End If
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
'           : connections (ODBC, etc.) via GetConnectivityProbeSql.
'           : Returns True if the server responds. Used to distinguish "server down"
'           : from "single table missing" when a linked table fails TableExists.
'---------------------------------------------------------------------------------------
'
Public Function TestBackEndConnection(strConnect As String) As Boolean

    Dim qdf As DAO.QueryDef
    Dim strPath As String
    Dim strKey As String
    Dim dbBackEnd As DAO.Database

    If IsAccessBackEndConnect(strConnect) Then
        strPath = GetConnectPath(strConnect)
        If Len(strPath) = 0 Then Exit Function

        strKey = UCase$(strPath)
        If m_dBackEndConnections Is Nothing Then Set m_dBackEndConnections = New Dictionary
        If m_dUnavailableBackEnds Is Nothing Then Set m_dUnavailableBackEnds = New Dictionary

        If m_dBackEndConnections.Exists(strKey) Then
            TestBackEndConnection = True
        ElseIf m_dUnavailableBackEnds.Exists(strKey) Then
            TestBackEndConnection = False
        Else
            Set dbBackEnd = OpenAccessBackEnd(strConnect)
            If Not dbBackEnd Is Nothing Then
                m_dBackEndConnections.Add strKey, dbBackEnd
                TestBackEndConnection = True
            End If
        End If
        Exit Function
    End If

    ' For ODBC and other connection types, attempt a lightweight query
    LogUnhandledErrors
    On Error Resume Next

    Set qdf = CurrentDb.CreateQueryDef("")
    qdf.Connect = strConnect
    qdf.SQL = GetConnectivityProbeSql(strConnect)
    qdf.OpenRecordset
    TestBackEndConnection = (Err.Number = 0)

    If Not qdf Is Nothing Then
        qdf.Close
        Set qdf = Nothing
    End If

    On Error GoTo 0

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetBackEndConnectKey
' Author    : Adam Waller
' Date      : 06/09/2026
' Purpose   : Public wrapper around GetBackEndKey for unit tests and diagnostics.
'---------------------------------------------------------------------------------------
'
Public Function GetBackEndConnectKey(strConnect As String) As String
    GetBackEndConnectKey = GetBackEndKey(strConnect)
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

    If IsAccessBackEndConnect(strConnect) Then
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
            GetConnectionEnvKey = ENV_KEY_PREFIX & Left$(GetStringHash(.GetStr), 7)
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
' Procedure : HasAuthInfo
' Author    : Adam Waller
' Date      : 06/19/2026
' Purpose   : Returns True if the connection string carries enough authentication
'           : detail to connect without prompting. Broader than HasCredentials: in
'           : addition to UID/PWD this recognizes trusted (SSPI) and explicit
'           : authentication methods such as Azure AD (Authentication=...), which
'           : have no password but still resolve the credential prompt.
'---------------------------------------------------------------------------------------
'
Private Function HasAuthInfo(strConnect As String) As Boolean
    HasAuthInfo = HasCredentials(strConnect) _
        Or (InStr(1, strConnect, "Authentication=", vbTextCompare) > 0) _
        Or (InStr(1, strConnect, "Trusted_Connection=", vbTextCompare) > 0) _
        Or (InStr(1, strConnect, "Integrated Security=", vbTextCompare) > 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : StripConnectionCredentials
' Author    : Adam Waller
' Date      : 06/19/2026
' Purpose   : Return the connection string with any UID= and PWD= segments removed.
'           : Used to ensure secrets are never written to source files.
'---------------------------------------------------------------------------------------
'
Public Function StripConnectionCredentials(strConnect As String) As String

    Dim varParts As Variant
    Dim lngPart As Long
    Dim strPart As String

    If Len(strConnect) = 0 Then Exit Function

    varParts = Split(strConnect, ";")
    With New clsConcat
        .AppendOnAdd = ";"
        For lngPart = 0 To UBound(varParts)
            strPart = CStr(varParts(lngPart))
            If Len(strPart) > 0 Then
                If Not StartsWith(strPart, "UID=", vbTextCompare) _
                    And Not StartsWith(strPart, "PWD=", vbTextCompare) Then
                    .Add strPart
                End If
            End If
        Next lngPart
        .Remove 1
        StripConnectionCredentials = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSourceSafeConnect
' Author    : Adam Waller
' Date      : 06/19/2026
' Purpose   : Return a connection string safe to write to a source file. Starting with
'           : export format v5.0.0, raw passwords are never written to source files in
'           : any UseEnvForConnections mode (including Never); credentials belong only
'           : in the git-ignored .env file. This removes the secret-leak footgun where
'           : a user (or an AI agent) commits source containing a plaintext password.
'           : A single warning is logged per distinct connection when stripping occurs.
'---------------------------------------------------------------------------------------
'
Public Function GetSourceSafeConnect(strConnect As String, strObjectLabel As String) As String

    GetSourceSafeConnect = strConnect
    If Len(strConnect) = 0 Then Exit Function

    ' Gated so existing repos that rely on self-contained source are not force-migrated.
    If Options.ExportFormatVersion < EFV_5_0_0 Then Exit Function

    ' Only act when an actual password value is present (avoids false positives on
    ' passwordless auth such as AD/integrated, which may carry an empty PWD=).
    If Len(GetConnectPart(strConnect, "PWD")) = 0 Then Exit Function

    GetSourceSafeConnect = StripConnectionCredentials(strConnect)

    ' Warn once per distinct connection so secrets are never silently shipped.
    If m_dStrippedConnWarn Is Nothing Then Set m_dStrippedConnWarn = New Dictionary
    If Not m_dStrippedConnWarn.Exists(GetSourceSafeConnect) Then
        m_dStrippedConnWarn.Add GetSourceSafeConnect, True
        Log.Error eelWarning, T("Removed credentials from the connection string for {0} " & _
            "to avoid storing secrets in source files. Store them in .env " & _
            "(UseEnvForConnections) or supply them at build time.", var0:=strObjectLabel), _
            ModuleName & ".GetSourceSafeConnect"
    End If

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

    Set m_cEnvResolved = Nothing

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

    Set cEnv = GetEnvCache(False)

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
'           : blnForWrite=True returns the base .env only (for export writes).
'           : blnForWrite=False returns the merged layered config (for resolution).
'---------------------------------------------------------------------------------------
'
Private Function GetEnvCache(blnForWrite As Boolean) As clsDotEnv

    Dim strFile As String

    If blnForWrite Then
        If m_cEnvCache Is Nothing Then
            Set m_cEnvCache = New clsDotEnv
            strFile = GetEnvFilePath
            If FSO.FileExists(strFile) Then
                m_cEnvCache.LoadFromFile strFile
            End If
        End If
        Set GetEnvCache = m_cEnvCache
    Else
        If m_cEnvResolved Is Nothing Then
            Set m_cEnvResolved = BuildResolvedEnv
        End If
        Set GetEnvCache = m_cEnvResolved
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildResolvedEnv
' Author    : Adam Waller
' Date      : 5/29/2026
' Purpose   : Build merged .env config: .env, .env.local, .env.{APP_ENV},
'           : .env.{APP_ENV}.local (later files override earlier).
'---------------------------------------------------------------------------------------
'
Private Function BuildResolvedEnv() As clsDotEnv

    Dim strDir As String
    Dim strAppEnv As String
    Dim cEnv As clsDotEnv

    strDir = Options.GetExportFolder
    Set cEnv = New clsDotEnv
    cEnv.LoadFromFileIfExists strDir & ".env"
    cEnv.LoadFromFileIfExists strDir & ".env.local", blnMerge:=True

    strAppEnv = Environ$(APP_ENV_KEY)
    If Len(strAppEnv) = 0 Then
        strAppEnv = cEnv.GetVar(APP_ENV_KEY, blnUseEnviron:=False)
    End If

    If Len(strAppEnv) > 0 Then
        cEnv.LoadFromFileIfExists strDir & ".env." & strAppEnv, blnMerge:=True
        cEnv.LoadFromFileIfExists strDir & ".env." & strAppEnv & ".local", blnMerge:=True
    End If

    Set BuildResolvedEnv = cEnv

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
    Set m_cEnvResolved = Nothing
    Set m_dEnvKeysWritten = Nothing
    Set m_dMissingEnvKeys = Nothing
    Set m_dStrippedConnWarn = Nothing
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

    Set cEnv = GetEnvCache(True)

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
