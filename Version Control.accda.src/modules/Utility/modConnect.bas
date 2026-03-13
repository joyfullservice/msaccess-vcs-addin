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
Private m_dCachedConnections As Dictionary
Private m_dBackEndConnections As Dictionary
Private m_dUnavailableBackEnds As Dictionary


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
Private Function GetConnectPart(strConnect As String, strPart As String) As String

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
