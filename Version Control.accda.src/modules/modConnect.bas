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

Private Const ModuleName = "modConnect"
Private m_dCachedConnections As Dictionary


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
                        Or Not Options.AggressiveSanitize Then
                        ' Retain the values if not using trusted connection, or if
                        ' AggressiveSanitize option is set to false (Defaults to true).
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
    Dim bolCached As Boolean

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
