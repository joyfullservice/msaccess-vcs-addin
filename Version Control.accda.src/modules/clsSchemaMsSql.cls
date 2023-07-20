VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSchemaMsSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This class extends the IDbSchema class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


' This requires us to use all the public methods and properties of the implemented class
' which keeps all the server classes consistent in how they are used in the export
' process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbSchema


' Handle local variables
Private Type udtThis
    blnInitialized As Boolean
    strName As String
    strBaseFolder As String
    strConnect As String
    strUserID As String
    strPassword As String
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_Export
' Author    : Adam Waller
' Date      : 7/14/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub IDbSchema_Export(blnFullExport As Boolean, Optional strAlternatePath As String)

    Dim dFile As Dictionary
    Dim dIndex As Dictionary
    Dim dObjects As Dictionary
    Dim dteLast As Date
    Dim rstObjects As ADODB.Recordset
    Dim blnUseGetDDL As Boolean
    Dim conn As ADODB.Connection
    Dim strItem As String
    Dim strPath As String
    Dim blnExport As Boolean
    Dim dItem As Dictionary

    ' Make sure we initialize before running the export
    If Not this.blnInitialized Then Exit Sub

    Perf.StartTiming

    ' Load index
    Set dFile = ReadJsonFile(IndexPath)
    If Not dFile Is Nothing Then
        Set dIndex = dFile("Items")
    Else
        Set dIndex = New Dictionary
    End If

    ' Return list of objects from the server
    If Not blnFullExport Then dteLast = Nz2(dNZ(dIndex, "Items\LastExport"), CDate(0))
    Set conn = New ADODB.Connection
    Perf.OperationStart "Retrieve SQL Objects"
    conn.Open this.strConnect, this.strUserID, this.strPassword
    Set rstObjects = conn.Execute(CodeDb.QueryDefs("qryMsSqlServerObjects").SQL)
    Perf.OperationEnd

    ' Loop through objects, building dictionary of items that match our filter.
    Perf.OperationStart "Loop through SQL objects"
    With rstObjects
        'strPath
        Do While Not .EOF

            ' Build item path and full path to source file
            strItem = Nz(!Folder) & PathSep & GetSafeFileName(Nz(!Name)) & ".sql"
            strPath = this.strBaseFolder & strItem

            ' Reset variables
            blnExport = True
            Set dItem = Nothing

            ' Check for item in index
            If dIndex.Exists(strItem) Then
                Set dItem = dIndex(strItem)
                ' Export if the dates are different
                blnExport = (dItem("LastModified") <> ConvertToIso((!last_modified)))
            End If

            ' Export if flag set
            If blnExport Then
                ExportObject Nz(!type_desc), Nz(!schema), Nz(!Name), strPath, conn, dIndex
                ' Update record in index
                UpdateIndex rstObjects, strItem, dIndex, dItem
            End If


            .MoveNext
        Loop
        .Close
    End With
    Perf.OperationEnd

    ' Close connection
    conn.Close

    Perf.EndTiming
    Debug.Print Perf.GetReports

    ' Now, loop back through the files and remove any file that is not represented by
    ' the list of objects returned from the server.


    ' Save the index
    WriteFile BuildJsonFile(TypeName(Me), dIndex, "Version Control System Schema Index"), this.strBaseFolder & "vcs-index.json"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportObject
' Author    : Adam Waller
' Date      : 7/18/2023
' Purpose   : Export the object definition and update the index
'---------------------------------------------------------------------------------------
'
Private Sub ExportObject(strType, strSchema As String, strName As String, strPath As String, _
    ByRef oConn As ADODB.Connection, ByRef dIndex As Dictionary)

    Static intUseSP As Integer

    Dim strSqlDef
    Dim strDefinition As String
    Dim rst As ADODB.Recordset
    Dim rst2 As ADODB.Recordset
    Dim strFullName As String
    Dim varTest As Variant
    Dim cmd As ADODB.Command
    Dim strCmdTemplate As String

    ' Prepare template statement for sp_GetDDL to work around VARCHAR(MAX) issue
    ' with many SQL Server ODBC drivers.
    strCmdTemplate = _
        "DECLARE @table TABLE (item text) " & _
        "INSERT INTO @table exec sp_GetDDL N'{name}' " & _
        "SELECT * FROM @table"

    ' Cache whether or not
    'if intusesp=0 then intusesp =

    strFullName = "[" & strSchema & "].[" & strName & "]"

    ' Determine how to export this type of object
    Select Case strType
        Case "USER_TABLE", "VIEW", "SYNONYM", "SQL_STORED_PROCEDURE", _
            "SQL_SCALAR_FUNCTION", "SQL_INLINE_TABLE_VALUED_FUNCTION", "SQL_TABLE_VALUED_FUNCTION"
            strSqlDef = Replace(strCmdTemplate, "{name}", strFullName)
            'strSqlDef = "exec sp_help " & strFullName & ""

        Case "TYPE_TABLE", "SEQUENCE_OBJECT", "SERVICE_QUEUE", "SYSTEM_TABLE", "INTERNAL_TABLE"
            ' Unsupported non-dependent objects

    End Select

    ' Sanity check
    If Len(strSqlDef) Then
        Perf.OperationStart "Run sp_GetDDL on " & strType
        Set cmd = New ADODB.Command
        With cmd
            Set .ActiveConnection = oConn
            .CommandText = strSqlDef
            Set rst = .Execute
        End With

        ' Get secondary recordset with object definition record
        Set rst2 = rst.NextRecordset
        With rst2
            If Not .EOF Then strDefinition = Nz(.Fields(0))
            If strDefinition = vbNullString Then
                If FSO.FileExists(strPath) Then DeleteFile strPath
            Else
                ' Export to file
                WriteFile strDefinition, strPath
            End If
            .Close
        End With
        Perf.OperationEnd
    End If

End Sub



'---------------------------------------------------------------------------------------
' Procedure : UpdateIndex
' Author    : Adam Waller
' Date      : 7/19/2023
' Purpose   : Update the entry in the index using values from the recordset
'---------------------------------------------------------------------------------------
'
Private Sub UpdateIndex(rst As ADODB.Recordset, strRelativePath As String, _
    ByRef dIndex As Dictionary, Optional ByRef dItem As Dictionary)

    Perf.OperationStart "Update Schema Index"
    ' See if we are working with an existing entry
    If dItem Is Nothing Then
        ' Add the new entry
        Set dItem = New Dictionary
        dIndex.Add strRelativePath, dItem
    End If

    ' Update the meta values
    dItem("ExportDate") = Now()
    dItem("LastModified") = Nz(rst!last_modified)
    Perf.OperationEnd

End Sub




Private Function PurgeOrphanedObjects()

End Function


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_Initialize
' Author    : Adam Waller
' Date      : 7/18/2023
' Purpose   : Initialize the database schema
'---------------------------------------------------------------------------------------
'
Private Sub IDbSchema_Initialize(dInstance As Scripting.IDictionary)
    With this
        .strName = dNZ(dInstance, "Name")
        .strConnect = dNZ(dInstance, "Connect")
        .strBaseFolder = Options.GetExportFolder & "databases\" & GetSafeFileName(.strName) & PathSep
        .blnInitialized = True
    End With
End Sub



Private Property Get IDbSchema_BaseFolder() As String
    IDbSchema_BaseFolder = this.strBaseFolder
End Property


Private Function IndexPath() As String
    IndexPath = this.strBaseFolder & "vcs-index.json"
End Function


Private Property Get IDbSchema_Name() As String

End Property


'---------------------------------------------------------------------------------------
' Procedure : IDbSchema_ServerType
' Author    : Adam Waller
' Date      : 7/14/2023
' Purpose   : Return server type
'---------------------------------------------------------------------------------------
'
Private Property Get IDbSchema_ServerType() As eDatabaseServerType
    IDbSchema_ServerType = estMsSql
End Property


Private Property Get IDbSchema_TypeDescription() As Variant
    IDbSchema_TypeDescription = "Microsoft SQL Server"
End Property


Private Property Get IDbSchema_ServerVersion() As String

End Property
