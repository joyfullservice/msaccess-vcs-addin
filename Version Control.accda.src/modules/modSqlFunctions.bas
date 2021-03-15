'---------------------------------------------------------------------------------------
' Module    : modAdpFunctions
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Functions for interacting with SQL servers and ADP projects.
'           : (ADP projects were discontinued after Microsoft Access 2010)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : GetSQLObjectModifiedDate
' Author    : Adam Waller
' Date      : 10/11/2017
' Purpose   : Get the last modified date for the SQL object
'---------------------------------------------------------------------------------------
'
Public Function GetSQLObjectModifiedDate(strName As String, eType As eSqlObjectType) As Date

    ' Use static variables so we can avoid hundreds of repeated calls
    ' for the same object type. Instead use a local array after
    ' pulling the initial values.
    ' (Makes a significant performance gain in complex databases)
    Static colCache As Collection
    Static strLastType As String
    Static dteCacheDate As Date

    Dim rst As ADODB.Recordset
    Dim strSql As String
    Dim strObject As String
    Dim strTypeFilter As String
    Dim intPos As Integer
    Dim strSchema As String
    Dim varItem As Variant
    Dim strType As String
    
    ' Shortcut to clear the cached variable
    If strName = vbNullString And strType = vbNullString Then
        Set colCache = Nothing
        strLastType = vbNullString
        dteCacheDate = 0
        Exit Function
    End If
    
    ' Only try this on ADP projects
    If CurrentProject.ProjectType <> acADP Then Exit Function
    
    ' Simple validation on object name
    strObject = Replace(strName, ";", vbNullString)
    
    ' Build schema filter if required
    intPos = InStr(1, strObject, ".")
    If intPos > 0 Then
        strObject = Mid$(strObject, intPos + 1)
        strSchema = Left$(strName, intPos - 1)
        'strSchemaFilter = " AND [schema_id]=schema_id('" & strSchema & "')"
    Else
        strSchema = "dbo"
    End If
    
    ' Build type filter
    Select Case eType
        Case estView: strType = "V"
        Case estStoredProcedure: strType = "P"
        Case estTable: strType = "U"
        Case estTrigger: strType = "TR"
    End Select
    If strType <> vbNullString Then strTypeFilter = " AND [type]='" & strType & "'"
    
    ' Check to see if we have already cached the results
    If strType = strLastType And (DateDiff("s", dteCacheDate, Now) < 5) And Not colCache Is Nothing Then
        ' Look through cache to find matching date
        For Each varItem In colCache
            If varItem(0) = strName Then
                GetSQLObjectModifiedDate = varItem(1)
                Exit For
            End If
        Next varItem
    Else
        ' Look up from query, and cache results
        Set colCache = New Collection
        dteCacheDate = Now
        strLastType = strType
        
        ' Build SQL query to find object
        strSql = "SELECT [name], schema_name([schema_id]) as [schema], modify_date FROM sys.objects WHERE 1=1 " & strTypeFilter
        Set rst = New ADODB.Recordset
        With rst
            .Open strSql, CurrentProject.Connection, adOpenForwardOnly, adLockReadOnly
            Do While Not .EOF
                ' Return date when name matches. (But continue caching additional results)
                If Nz(!Name) = strObject And Nz(!schema) = strSchema Then GetSQLObjectModifiedDate = Nz(!modify_date)
                If Nz(!schema) = "dbo" Then
                    colCache.Add Array(Nz(!Name), Nz(!modify_date))
                Else
                    ' Include schema name in object name
                    colCache.Add Array(Nz(!schema) & "." & Nz(!Name), Nz(!modify_date))
                End If
                .MoveNext
            Loop
            .Close
        End With
        Set rst = Nothing
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSQLObjectDefinitionForADP
' Author    : awaller
' Date      : 12/12/2016
' Purpose   : Returns the SQL definition for the ADP project item.
'           : (Queries, Views, Tables, etc... are not stored in Access but on the
'           :  SQL server.)
'           : NOTE: This takes a simplistic approach, which does not guard againts
'           : certain types of SQL injection attacks. Use at your own risk!
'---------------------------------------------------------------------------------------
'
Public Function GetSQLObjectDefinitionForADP(strName As String) As String
    
    Dim rst As ADODB.Recordset
    Dim strSql As String
    Dim strObject As String
    
    ' Only try this on ADP projects
    If CurrentProject.ProjectType <> acADP Then Exit Function
    
    ' Simple validation on object name
    strObject = Replace(strName, ";", vbNullString)
    
    strSql = "SELECT object_definition (OBJECT_ID(N'" & strObject & "'))"
    '@Ignore SetAssignmentWithIncompatibleObjectType
    Set rst = CurrentProject.Connection.Execute(strSql)
    If Not rst.EOF Then
        ' Get SQL definition
        GetSQLObjectDefinitionForADP = Nz(rst(0).Value)
    End If
    
    Set rst = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : StripDboPrefix
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Removes the dbo prefix, as sometimes encountered with ADP projects
'           : depending on the sql permissions of the current user.
'---------------------------------------------------------------------------------------
'
Public Function StripDboPrefix(strName As String) As String
    If Left$(strName, 4) = "dbo." Then
        StripDboPrefix = Mid$(strName, 5)
    Else
        StripDboPrefix = strName
    End If
End Function



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
        
        ' Remove trailing semicolon, and return string
        .Remove 1
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
Private Function GetFullConnect(strRelative As String) As String

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