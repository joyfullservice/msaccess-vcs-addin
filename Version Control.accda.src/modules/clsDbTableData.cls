VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbTableData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This class extends the IDbComponent class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Public Format As eTableDataExportFormat

Const UNSUPPORTED_DATA_TYPE As String = "UNSUPPORTED DATA TYPE"

Private m_Table As AccessObject
Private m_Items(True To False) As Dictionary
Private m_FileList As Dictionary

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export(Optional strAlternatePath As String)

    Dim strFile As String
    Dim intFormat As eTableDataExportFormat

    ' Use error handling in case we encouter errors during export
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Make sure the table exists, and is accessible
    If Not TableExists(m_Table.Name) Then
        Log.Error eelError, "Unable to access table " & m_Table.Name & _
            ". Please ensure that this table exists and is accessible in the current database.", _
            ModuleName(Me) & ".Export"
        Exit Sub
    End If

    ' Save as selected format, and remove other formats if they exist.
    For intFormat = 1 To eTableDataExportFormat.[_Last]
        ' Build file name for this format
        strFile = Nz2(strAlternatePath, IDbComponent_BaseFolder & GetSafeFileName(m_Table.Name) & "." & GetExtByFormat(intFormat))
        If FSO.FileExists(strFile) Then DeleteFile strFile, True
        If intFormat = Me.Format Then
            ' Export the table using this format.
            Select Case intFormat
                Case etdTabDelimited
                    ' Export in tab delimited text file
                    ExportTableDataAsTDF m_Table.Name, strFile
                Case etdXML
                    ' Export data rows as XML (encoding default is UTF-8)
                    VerifyPath strFile
                    Perf.OperationStart "App.ExportXML()"
                    If IsLocalTable(m_Table.Name) Then
                        ' Embedded schema for local tables to support calculated fields
                        Application.ExportXML acExportTable, m_Table.Name, strFile, , , , , acEmbedSchema
                    Else
                        ' Export data only for linked tables
                        Application.ExportXML acExportTable, m_Table.Name, strFile
                    End If
                    Perf.OperationEnd
                    With New clsSourceParser
                        .LoadSourceFile strFile
                        WriteFile .Sanitize(ectXML), strFile
                    End With
            End Select
        End If
    Next intFormat

    ' Trap any errors while exporting table data.
    CatchAny eelError, "Error exporting table data for " & m_Table.Name, ModuleName(Me) & ".Export"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportTableData
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Export the data from the table.
'---------------------------------------------------------------------------------------
'
Private Sub ExportTableDataAsTDF(strTable As String, strFile As String)

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim fld As DAO.Field
    Dim cData As New clsConcat
    Dim intFields As Integer
    Dim intCnt As Integer
    Dim lngProgMax As Long
    Dim lngProgValue As Long

    Perf.OperationStart "Export Table Data as TDF"

    ' Open table in fast read-only view
    Set dbs = CurrentDb
    Set rst = dbs.OpenRecordset(GetTableExportSql(strTable), dbOpenSnapshot, dbReadOnly)
    intFields = rst.Fields.Count

    ' Add header row
    For Each fld In rst.Fields
        cData.Add fld.Name
        intCnt = intCnt + 1
        If intCnt < intFields Then cData.Add vbTab
    Next fld
    cData.Add vbCrLf

    ' Make sure we actually have records
    If Not rst.EOF Then

        ' Stash current progress bar values so we can reuse the
        ' progress bar to show the progress through the rows.
        With Log.ProgressBar
            lngProgMax = .Max
            lngProgValue = .Value
            ' Set to recordset after accessing all records
            rst.MoveLast
            .Max = rst.RecordCount
            .Value = 1
            rst.MoveFirst
        End With

        ' Add data rows
        Do While Not rst.EOF
            intCnt = 0
            For Each fld In rst.Fields
                If IsBinaryType(fld) Then
                    ' Binary data not supported
                    cData.Add UNSUPPORTED_DATA_TYPE
                Else
                    ' Format for TDF format without line breaks
                    cData.Add FormatStringForTDF(Nz(fld.Value))
                End If
                intCnt = intCnt + 1
                If intCnt < intFields Then cData.Add vbTab
            Next fld
            cData.Add vbCrLf
            rst.MoveNext
            Log.Increment ' Increment log, in case this takes a while
        Loop

        ' Restore progress bar values for object iteration
        With Log.ProgressBar
             .Max = lngProgMax
             .Value = lngProgValue
        End With
    End If

    ' Close database objects
    rst.Close
    Set rst = Nothing

    Perf.OperationEnd

    ' Save output file
    WriteFile cData.GetStr, strFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportTableDataTDF
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Imports the data from a TDF file, loading it into the table
'---------------------------------------------------------------------------------------
'
Private Sub ImportTableDataTDF(strFile As String)

    Dim strTable As String
    Dim dCols As Dictionary
    Dim fld As DAO.Field
    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim stm As ADODB.Stream
    Dim strLine As String
    Dim varLine As Variant
    Dim varHeader As Variant
    Dim intCol As Integer
    Dim strValue As String

    Perf.OperationStart "Import TDF Data"

    ' Build a dictionary of column names so we can load the data
    ' into the matching columns.
    strTable = GetObjectNameFromFileName(strFile)
    Set dbs = CurrentDb
    Set dCols = New Dictionary
    For Each fld In dbs.TableDefs(strTable).Fields
        dCols.Add fld.Name, fld.Name
    Next fld

    ' Clear any existing records before importing this data.
    dbs.Execute "delete from [" & strTable & "]", dbFailOnError
    Set rst = dbs.OpenRecordset(strTable)

    ' Read file line by line
    Set stm = New ADODB.Stream
    With stm
        .Charset = "utf-8"
        .Open
        .LoadFromFile strFile
    End With

    ' Loop through lines in file
    Do While Not stm.EOS
        strLine = stm.ReadText(adReadLine)
        ' See if the header has already been parsed.
        If Not IsArray(varHeader) Then
            ' Skip past any UTF-8 BOM header
            If Left$(strLine, 3) = UTF8_BOM Then strLine = Mid$(strLine, 4)
            ' Read header line
            varHeader = Split(strLine, vbTab)
        Else
            ' Data line
            varLine = Split(strLine, vbTab)
            rst.AddNew
                ' Loop through fields
                For intCol = 0 To UBound(varHeader)
                    ' Check to see if field exists in the table
                    If dCols.Exists(varHeader(intCol)) Then
                        ' Check for empty string or null.
                        If varLine(intCol) = vbNullString Then
                            With rst.Fields(varHeader(intCol))
                                If IsNull(.Value) Then
                                    ' Could possibly hit a problem with the storage of
                                    ' zero length strings instead of nulls. Since we can't
                                    ' really differentiate between these in a TDF file,
                                    ' we will do some additional probing...
                                    If .AllowZeroLength And .Required Then
                                        ' Use empty string instead of null value
                                        .Value = vbNullString
                                    End If
                                End If
                            End With
                        Else
                            ' Perform any needed replacements
                            strValue = FormatStringFromTDF(CStr(varLine(intCol)))
                            If strValue <> CStr(varLine(intCol)) Then
                                ' Use replaced string value
                                rst.Fields(varHeader(intCol)).Value = strValue
                            Else
                                If strValue <> UNSUPPORTED_DATA_TYPE Then
                                    ' Use variant value without the string conversion
                                    rst.Fields(varHeader(intCol)).Value = varLine(intCol)
                                End If
                            End If
                        End If
                    End If
                Next intCol
            rst.Update
        End If
        ' Increment log, just in case this takes a while.
        Log.Increment
    Loop
    stm.Close
    Set stm = Nothing
    rst.Close
    Set rst = Nothing

    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : FormatStringForTDF
' Author    : Adam Waller
' Date      : 7/16/2021
' Purpose   : Replace line feeds and similar characters with escaped codes for
'           : representation in tab-delimited format.
'           : (Using Chr(26) as interim placeholder) See #251
'---------------------------------------------------------------------------------------
'
Private Function FormatStringForTDF(strValue As String) As String
    FormatStringForTDF = MultiReplace(strValue, _
        "\", Chr$(26), _
        vbCrLf, "\r\n", _
        vbCr, "\r", _
        vbLf, "\n", _
        vbTab, "\t", _
        Chr$(26), "\\")
End Function


'---------------------------------------------------------------------------------------
' Procedure : FormatStringFromTDF
' Author    : Adam Waller
' Date      : 7/16/2021
' Purpose   : Restore original characters from escaped codes.
'---------------------------------------------------------------------------------------
'
Private Function FormatStringFromTDF(strTDFValue) As String
    FormatStringFromTDF = MultiReplace(strTDFValue, _
        "\\", Chr$(26), _
        "\r\n", vbCrLf, _
        "\r", vbCr, _
        "\n", vbLf, _
        "\t", vbTab, _
        Chr$(26), "\")
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTableExportSql
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Build SQL to export `tbl_name` sorted by each field from first to last
'---------------------------------------------------------------------------------------
'
Private Function GetTableExportSql(strTable As String) As String

    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim intFields As Integer
    Dim cText As New clsConcat
    Dim cFieldList As New clsConcat
    Dim cSortList As New clsConcat
    Dim dbs As Database

    Set dbs = CurrentDb
    Set tdf = dbs.TableDefs(strTable)
    intFields = tdf.Fields.Count

    ' Build list of fields
    With cFieldList
        For Each fld In tdf.Fields
            .Add "[", fld.Name, "], "
            If Not IsBinaryType(fld) Then cSortList.Add "[", fld.Name, "], "
        Next fld
    End With

    ' Remove any trailing commas
    cFieldList.Remove 2
    cSortList.Remove 2

    ' Build select statement
    With cText
        .Add "SELECT ", cFieldList.GetStr
        .Add " FROM [", strTable, "] ORDER BY "
        .Add cSortList.GetStr
    End With

    GetTableExportSql = cText.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller, Florian Jenn
' Date      : 4/23/2020, 2020-10-26
' Purpose   : Import the table data from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim blnUseTemp As Boolean
    Dim strTempFile As String
    Dim strTable As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Make sure table exists before importing data to it.
    strTable = GetObjectNameFromFileName(strFile)

    ' Only allow import into local tables (not linked tables)
    If IsLocalTable(strTable) Then

        ' Import from different formats (XML is preferred for data integrity)
        Select Case GetFormatByExt(strFile)
            Case etdXML

                    ' The ImportXML function does not properly handle UrlEncoded paths
                    blnUseTemp = (InStr(1, strFile, "%") > 0)
                    If blnUseTemp Then
                        ' Import from (safe) temporary file name.
                        strTempFile = GetTempFile
                        FSO.CopyFile strFile, strTempFile
                        Application.ImportXML strTempFile, acAppendData
                        DeleteFile strTempFile
                    Else
                        Application.ImportXML strFile, acAppendData
                    End If

                    ' Trap any XML import errors
                    CatchAny eelError, "Error importing XML data from '" & strFile & "'", ModuleName(Me) & ".Import"

            Case etdTabDelimited
                ImportTableDataTDF strFile
        End Select

    Else
        ' Either not a local table, or it doesn't exist
        If TableExists(strTable) Then
            Log.Error eelWarning, "Table data may only be imported into local tables", "clsDbTableData.Import"
        Else
            ' Warn user that table does not exist.
            Log.Error eelError, "Table structure not found for '" & strTable & "'.", "clsDbTableData.Import"
            Log.Add "Table definition does not exist for '" & strTable & _
                "'. This must be created before importing table data.", False
        End If
        Log.Add "Table data from '" & strTable & "' was not imported.", False
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Merge
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Merge the source file into the existing database, updating or replacing
'           : any existing object.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Merge(strFile As String)
    Log.Error eelWarning, "Merge not supported for table data. (" & _
        GetObjectNameFromFileName(strFile) & ")", "clsDbTableData.Merge"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IDbComponent_MoveSource
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Move the component's source file(s) from one folder to another
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_MoveSource(strFromFolder As String, strToFolder As String)
    MoveFileIfExists strFromFolder & FSO.GetFileName(IDbComponent_SourceFile), strToFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary

    Dim tbl As AccessObject
    Dim cTable As clsDbTableData
    Dim blnAdd As Boolean

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        Set m_Items(blnModifiedOnly) = New Dictionary

        ' No need to go any further if we don't have any saved tables defined
        If Options.TablesToExportData.Count > 0 Then

            ' We have at least one table defined. Loop through the tables looking
            ' for a matching name.
            With Options
                For Each tbl In CurrentData.AllTables
                    If .TablesToExportData.Exists(tbl.Name) Then
                        Set cTable = New clsDbTableData
                        cTable.Format = .GetTableExportFormat(CStr(.TablesToExportData(tbl.Name)("Format")))
                        Set cTable.Parent.DbObject = tbl
                        blnAdd = True
                        If blnModifiedOnly Then blnAdd = cTable.IDbComponent_IsModified
                        If blnAdd Then m_Items(blnModifiedOnly).Add cTable.Parent.SourceFile, cTable
                    End If
                Next tbl
            End With
        End If
    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_Items(blnModifiedOnly)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type. (Could be
'           : a couple different file extensions involved.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Dictionary
    If m_FileList Is Nothing Then
        Set m_FileList = GetFilePathsInFolder(IDbComponent_BaseFolder, "*." & GetExtByFormat(etdTabDelimited))
        MergeDictionary m_FileList, GetFilePathsInFolder(IDbComponent_BaseFolder, "*." & GetExtByFormat(etdXML))
    End If
    Set IDbComponent_GetFileList = m_FileList
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsBinaryType
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Returns true if the field type uses binary content. (Export/import on
'           : binary data through XML or text is not well supported.)
'---------------------------------------------------------------------------------------
'
Private Function IsBinaryType(fld As DAO.Field) As Boolean
    Select Case fld.Type
        Case dbLongBinary, dbVarBinary, dbAttachment: IsBinaryType = True
        Case Else: IsBinaryType = False
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetExtByFormat
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the expected file extension by format.
'---------------------------------------------------------------------------------------
'
Private Function GetExtByFormat(intFormat As eTableDataExportFormat) As String
    Select Case intFormat
        Case etdTabDelimited:   GetExtByFormat = "txt"
        Case etdXML:            GetExtByFormat = "xml"
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFormatByExt
' Author    : Adam Waller
' Date      : 5/7/2020
' Purpose   : Look up the format from the extension name
'---------------------------------------------------------------------------------------
'
Private Function GetFormatByExt(strFile As String) As eTableDataExportFormat
    Select Case FSO.GetExtensionName(strFile)
        Case "txt": GetFormatByExt = etdTabDelimited
        Case "xml": GetFormatByExt = etdXML
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsModified
' Author    : Adam Waller
' Date      : 11/21/2020
' Purpose   : Returns true if the object in the database has been modified since
'           : the last export of the object.
'---------------------------------------------------------------------------------------
'
Public Function IDbComponent_IsModified() As Boolean
    ' It's not really worth it to try to determine if the table data has been changed.
    ' We would have to fully export anyway just to check, so let's always export for
    ' table data. (This is handled different from just about every other object type.)
    IDbComponent_IsModified = True
End Function


'---------------------------------------------------------------------------------------
' Procedure : DateModified
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'           : If the modified date cannot be determined (such as application
'           : properties) then this function will return 0.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_DateModified() As Date
    ' We cannot determine when *records* were modified in a table.
    IDbComponent_DateModified = 0
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "Table Data"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "tables" & PathSep
End Property


'---------------------------------------------------------------------------------------
' Procedure : FileExtensions
' Author    : Adam Waller
' Date      : 12/1/2023
' Purpose   : A collection of the file extensions used in source files for this
'           : component type.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_FileExtensions() As Collection
    Set IDbComponent_FileExtensions = New Collection
    IDbComponent_FileExtensions.Add "txt"
    IDbComponent_FileExtensions.Add "tdf"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    If m_Table Is Nothing Then Exit Property
    IDbComponent_Name = m_Table.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    If m_Table Is Nothing Then Exit Property
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_Table.Name) & "." & GetExtByFormat(Me.Format)
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count(Optional blnModifiedOnly As Boolean = False) As Long
    IDbComponent_Count = IDbComponent_GetAllFromDB(blnModifiedOnly).Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : QuickCount
' Author    : Adam Waller
' Date      : 6/14/2022
' Purpose   : Return a cached, non-iterative approximate count of database objects
'           : for use with progress indicators when scanning for changes. Single file
'           : objects like database properties can simply return 1.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_QuickCount() As Long
    ' In most cases very few tables will be involved here
    IDbComponent_QuickCount = 1
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbTableData
End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Returns true if the export of all items is done as a single file instead
'           : of individual files for each component. (I.e. properties, references)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SingleFile() As Boolean
    IDbComponent_SingleFile = False
End Property


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_Table
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Table = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : Parent
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Return a reference to this class as an IDbComponent. This allows you
'           : to reference the public methods of the parent class without needing
'           : to create a new class object.
'---------------------------------------------------------------------------------------
'
Public Property Get Parent() As IDbComponent
    Set Parent = Me
End Property
