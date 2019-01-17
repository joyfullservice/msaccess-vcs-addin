Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : GetSQLObjectModifiedDate
' Author    : Adam Waller
' Date      : 10/11/2017
' Purpose   : Get the last modified date for the SQL object
'---------------------------------------------------------------------------------------
'
Public Function GetSQLObjectModifiedDate(strName As String, ByVal strType As String) As Date

    ' Use static variables so we can avoid hundreds of repeated calls
    ' for the same object type. Instead use a local array after
    ' pulling the initial values.
    ' (Makes a significant performance gain in complex databases)
    Static colCache As Collection
    Static strLastType As String
    Static dteCacheDate As Date

    Dim rst As ADODB.Recordset
    Dim strSQL As String
    Dim strObject As String
    Dim strTypeFilter As String
    Dim intPos As Integer
    Dim strSchema As String
    Dim strSchemaFilter As String
    Dim varItem As Variant
    
    ' Shortcut to clear the cached variable
    If strName = "" And strType = "" Then
        Set colCache = Nothing
        strLastType = ""
        dteCacheDate = 0
        Exit Function
    End If
    
    ' Only try this on ADP projects
    If CurrentProject.ProjectType <> acADP Then Exit Function
    
    ' Simple validation on object name
    strObject = Replace(strName, ";", "")
    
    ' Build schema filter if required
    intPos = InStr(1, strObject, ".")
    If intPos > 0 Then
        strObject = Mid(strObject, intPos + 1)
        strSchema = Left(strName, intPos - 1)
        'strSchemaFilter = " AND [schema_id]=schema_id('" & strSchema & "')"
    Else
        strSchema = "dbo"
    End If
    
    ' Build type filter
    Select Case strType
        Case "V", "VIEW", "views": strType = "V"
        Case "P", "SQL_STORED_PROCEDURE", "procedures": strType = "P"
        Case "T", "TABLE", "U", "USER_TABLE", "tables": strType = "U"
        Case "TR", "SQL_TRIGGER", "triggers": strType = "TR"
        Case Else
            strType = strType
    End Select
    If strType <> "" Then strTypeFilter = " AND [type]='" & strType & "'"
    
    ' Check to see if we have already cached the results
    If strType = strLastType And (DateDiff("s", dteCacheDate, Now()) < 5) And Not colCache Is Nothing Then
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
        dteCacheDate = Now()
        strLastType = strType
        
        ' Build SQL query to find object
        strSQL = "SELECT [name], schema_name([schema_id]) as [schema], modify_date FROM sys.objects WHERE 1=1 " & strTypeFilter
        Set rst = New ADODB.Recordset
        With rst
            .Open strSQL, CurrentProject.Connection, adOpenForwardOnly, adLockReadOnly
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
    Dim strSQL As String
    Dim strObject As String
    
    ' Only try this on ADP projects
    If CurrentProject.ProjectType <> acADP Then Exit Function
    
    ' Simple validation on object name
    strObject = Replace(strName, ";", "")
    
    strSQL = "SELECT object_definition (OBJECT_ID(N'" & strObject & "'))"
    Set rst = CurrentProject.Connection.Execute(strSQL)
    If Not rst.EOF Then
        ' Get SQL definition
        GetSQLObjectDefinitionForADP = Nz(rst(0).Value)
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetADPTableDef
' Author    : awaller
' Date      : 12/12/2016
' Purpose   : Get the definition for an ADP table from SQL
'---------------------------------------------------------------------------------------
'
Public Function GetADPTableDef(strTable As String) As String

    Dim rst As ADODB.Recordset
    Dim strSQL As String
    Dim strObject As String
    Dim intRst As Integer
    Dim strText As String
    Dim fld As ADODB.Field
    Dim colText As New clsConcat
    
    ' Initialize counter
    intRst = 2
    
    ' Only try this on ADP projects
    If CurrentProject.ProjectType <> acADP Then Exit Function
    
    ' Simple validation on object name
    strObject = Replace(strTable, ";", "")
    
    ' Get initial table information
    strSQL = "exec sp_help N'" & strObject & "'"
    Set rst = CurrentProject.Connection.Execute(strSQL)
    colText.Add "-- sp_help Recordset 1" & vbCrLf & vbCrLf
    For Each fld In rst.Fields
        colText.Add fld.Name
        colText.Add vbTab
    Next fld
    colText.Add vbCrLf
    colText.Add rst.GetString(, , vbTab, vbCrLf)
    
    ' Loop through additional recordsets for columns, keys and other data
    Do
        Set rst = rst.NextRecordset
        If rst Is Nothing Then Exit Do
        If rst.State = adStateClosed Then Exit Do
        
        colText.Add vbCrLf & vbCrLf & "-- sp_help Recordset " & intRst & vbCrLf & vbCrLf
        For Each fld In rst.Fields
            colText.Add fld.Name
            colText.Add vbTab
        Next fld
        If Not rst.EOF Then
            colText.Add vbCrLf
            colText.Add rst.GetString(, , vbTab, vbCrLf)
        End If
        
        intRst = intRst + 1
    Loop
    
    ' Clear references
    Set fld = Nothing
    Set rst = Nothing
    
    GetADPTableDef = colText.GetStr
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExportADPTriggers
' Author    : awaller
' Date      : 12/14/2016
' Purpose   : Export the triggers
'---------------------------------------------------------------------------------------
'
Public Sub ExportADPTriggers(cModel As IVersionControl, strBaseExportFolder As String)

    Dim colTriggers As New Collection
    Dim rst As ADODB.Recordset
    Dim strSQL As String
    Dim strDef As String
    Dim strFile As String
    Dim varTrg As Variant
    Dim blnFound As Boolean
    Dim dteFileModified As Date
    Dim blnSkip As Boolean
    
    ' Only try this on ADP projects
    If CurrentProject.ProjectType <> acADP Then Exit Sub
    
    ' Build list of triggers in database (from sysobjects)
    strSQL = "SELECT [name],object_name(parent_object_id) AS parent_name, schema_name([schema_id]) AS [schema_name], modify_date FROM sys.objects WHERE type='TR'"
    Set rst = New ADODB.Recordset
    With rst
        .Open strSQL, CurrentProject.Connection, adOpenForwardOnly, adLockReadOnly
        Do While Not .EOF
            strFile = GetSafeFileName(Nz(!schema_name) & "_" & Nz(!Name) & ".sql")
            colTriggers.Add Array(Nz(!Name), Nz(!parent_name), Nz(!schema_name), strFile, Nz(!modify_date))
            .MoveNext
        Loop
        .Close
    End With
    Set rst = Nothing
    
    ' If no triggers, then clear and exit
    If colTriggers.Count = 0 Then
        If DirExists(strBaseExportFolder) Then
            ClearTextFilesFromDir strBaseExportFolder, "sql"
            Exit Sub
        End If
    End If
    
    ' Prepare folder
    If Not DirExists(strBaseExportFolder) Then VerifyPath strBaseExportFolder
    
    
    ' Clear all existing files unless we are using fast save.
    If cModel.FastSave Then
    
        ' Loop through saved source files, removing ones that no longer exist in the database.
        strFile = Dir(strBaseExportFolder & "*.sql")
        Do While strFile <> ""
            blnFound = False
            For Each varTrg In colTriggers
                If varTrg(3) = strFile Then
                    ' Found matching object in database
                    blnFound = True
                    Exit For
                End If
            Next varTrg
            If Not blnFound Then
                ' No matching object found
                Kill strBaseExportFolder & strFile
            End If
            strFile = Dir()
        Loop
    Else
        ' Not using fast save.
        ClearTextFilesFromDir strBaseExportFolder, "sql"
    End If
    

    ' Now go through and export the triggers
    For Each varTrg In colTriggers
        
        ' Check for fast save, to see if we can just export the newly changed triggers
        If cModel.FastSave Then
            strFile = strBaseExportFolder & varTrg(3)
            If Not FileExists(strFile) Then
                blnSkip = False
            Else
                dteFileModified = FileDateTime(strFile)
                If varTrg(4) > dteFileModified Then
                    ' Changed in SQL server
                    blnSkip = False
                Else
                    ' Appears unchanged from the modified dates
                    blnSkip = True
                End If
            End If
        End If
        
        If blnSkip Then
            If ShowDebugInfo Then Debug.Print "    (Skipping) [Trigger] - " & varTrg(0)
        Else
            ' Export the trigger definition
            strDef = GetSQLObjectDefinitionForADP(varTrg(2) & "." & varTrg(0))
            WriteFile strDef, strBaseExportFolder & varTrg(3)
            ' Show output
            If ShowDebugInfo Then Debug.Print "    [Trigger] - " & varTrg(0)
        End If
            
    Next varTrg

    
End Sub