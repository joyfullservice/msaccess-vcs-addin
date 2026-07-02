Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSOptionsTableData
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Tables & Data options section: table data export selection with filters.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Implements IOptionsSection

Private Enum eTableCol
    etcName = 0
    etcType = 1
    etcHidden = 2
    etcSystem = 3
    etcOther = 4
    etcLocal = 5
    etcLinked = 6
End Enum


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Load table list and populate format combo boxes.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    Dim intFormat As eTableDataExportFormat
    Dim opt As clsOptions

    Set opt = Form_frmVCSOptions.OptionsInstance

    ' Populate format type combo boxes
    Dim frmTableData As Form_frmVCSTableData
    Set frmTableData = Me.sfrmTableData.Form
    With frmTableData.cboFormatType
        .RowSource = vbNullString
        For intFormat = eTableDataExportFormat.etdNoData To eTableDataExportFormat.[_Last] - 1
            .AddItem intFormat & ";""" & opt.GetTableExportFormatName(intFormat) & """", intFormat
        Next intFormat
        Me.cboFormatTypeForOther.RowSource = .RowSource
        Me.cboFormatTypeForOther.RemoveItem etdNoData
    End With

    LoadOptions

    ' Set initial column sizing
    DoEvents
    Form_frmVCSTableData.Form_Resize

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IOptionsSection_LoadOptions / IOptionsSection_SaveOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : IOptionsSection interface implementation.
'---------------------------------------------------------------------------------------
'
Private Sub IOptionsSection_LoadOptions()
    LoadOptions
End Sub

Private Sub IOptionsSection_SaveOptions()
    SaveOptions
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Load the table list from the working options instance.
'---------------------------------------------------------------------------------------
'
Private Sub LoadOptions()
    LoadTableList
    RefreshTableDisplay
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Save the table data selections back to the working options instance.
'---------------------------------------------------------------------------------------
'
Private Sub SaveOptions()
    SaveTableList
End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkTableShowHidden_Click / chkTableShowOther_Click / chkTableShowSystem_Click
' Author    : Adam Waller
' Date      : 4/17/2020
' Purpose   : Update the list of tables when filter checkboxes change.
'---------------------------------------------------------------------------------------
'
Private Sub chkTableShowHidden_Click()
    RefreshTableDisplay
End Sub

Private Sub chkTableShowOther_Click()
    RefreshTableDisplay
End Sub

Private Sub chkTableShowSystem_Click()
    RefreshTableDisplay
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdAddOtherTableData_Click
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Add a table not in the current database to the export list.
'---------------------------------------------------------------------------------------
'
Private Sub cmdAddOtherTableData_Click()
    Dim strTable As String
    If Nz(Me.cboFormatTypeForOther.Value, "No Data") = T("No Data") Then
        MsgBox2 T("Please select format"), T("Select the format to save table data before adding the table to the list."), , vbInformation
    Else
        strTable = Nz(Me.txtOtherTableName.Value, vbNullString)
        If strTable <> vbNullString Then
            AddUpdateTableInList strTable, Nz(Me.cboFormatTypeForOther.Value), False, False, True, False
            chkTableShowOther = True
            RefreshTableDisplay
        End If
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadTableList
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Load the list of tables in the current database, and merge in the list
'           : of tables where we are opting to save data.
'---------------------------------------------------------------------------------------
'
Private Sub LoadTableList()

    Dim varKey As Variant
    Dim strFormat As String
    Dim intFormat As eTableDataExportFormat
    Dim strName As String
    Dim dbs As DAO.Database
    Dim rstTableData As DAO.Recordset
    Dim rstSource As DAO.Recordset
    Dim strSql As String
    Dim lngFlags As Long
    Dim lngType As Long
    Dim opt As clsOptions

    Set opt = Form_frmVCSOptions.OptionsInstance

    ' Reset list of tables
    Set dbs = CodeDb
    dbs.Execute "DELETE FROM tblTableData;", dbFailOnError

    ' Open table to load records
    Set rstTableData = dbs.OpenRecordset("SELECT * FROM tblTableData;", dbOpenDynaset)

    ' Get list of tables if we have a database file open.
    If DatabaseFileOpen Then
        If Not (CurrentProject.ProjectType = acADP) Then
            strSql = _
                "SELECT o.Name, o.Type, o.Flags " & _
                "FROM MSysObjects AS o " & _
                "WHERE o.Type IN (1, 4, 6) " & _
                "ORDER BY o.Name;"

            Set rstSource = CurrentDb.OpenRecordset(strSql, dbOpenSnapshot)
            With rstSource
                Do While Not .EOF
                    lngFlags = Nz(!Flags, 0)
                    lngType = Nz(!Type, 0)
                    If (lngFlags < 0) Or BitSet(lngFlags, 1) Then
                        ' Don't include read-only or deeply hidden system tables.
                    Else
                        rstTableData.AddNew
                            rstTableData!TableName = Nz(!Name)
                            rstTableData!Flags = Nz(!Flags)
                            rstTableData!IsSystem = BitSet(lngFlags, 2)
                            rstTableData!IsHidden = BitSet(lngFlags, 8)
                            rstTableData!IsLocal = (lngType = 1)
                            rstTableData!TableIcon = GetTableIcon(etcLinked)
                            If rstTableData!IsLocal Then rstTableData!TableIcon = GetTableIcon(etcLocal)
                            If rstTableData!IsSystem Then rstTableData!TableIcon = GetTableIcon(etcSystem)
                        rstTableData.Update
                    End If
                    .MoveNext
                Loop
                .Close
            End With
        End If
    End If

    ' Merge in tables from saved options
    If Not opt.TablesToExportData Is Nothing Then
        For Each varKey In opt.TablesToExportData.Keys
            strName = CStr(varKey)
            strFormat = opt.TablesToExportData.Item(varKey)("Format")
            intFormat = opt.GetTableExportFormat(strFormat)

            With rstTableData
                .FindFirst "[TableName]='" & Replace$(strName, "'", "''") & "'"
                If .NoMatch Then
                    .AddNew
                    !TableName = strName
                    !TableIcon = GetTableIcon(etcOther)
                    !FormatType = intFormat
                    !IsOther = True
                    .Update
                Else
                    .Edit
                    !FormatType = intFormat
                    .Update
                End If
            End With
        Next varKey
    End If

    rstTableData.Close

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveTableList
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Save the list of tables that should have their data exported.
'---------------------------------------------------------------------------------------
'
Private Sub SaveTableList()

    Dim rstTableData As DAO.Recordset
    Dim dTables As Dictionary
    Dim dTable As Dictionary
    Dim opt As clsOptions

    Set opt = Form_frmVCSOptions.OptionsInstance

    Set dTables = New Dictionary
    dTables.CompareMode = TextCompare

    If Not (CurrentProject.ProjectType = acADP) Then
        Set rstTableData = CodeDb.OpenRecordset( _
            "SELECT TableName, FormatType FROM tblTableData " & _
            "WHERE FormatType <> 0 ORDER BY TableName;", dbOpenForwardOnly)
        With rstTableData
            Do Until .EOF
                Set dTable = New Dictionary
                dTable.CompareMode = TextCompare
                dTable("Format") = opt.GetTableExportFormatName(Nz(!FormatType, 0))
                dTables.Add Nz(!TableName), dTable
                .MoveNext
            Loop
            .Close
        End With
    End If

    Set opt.TablesToExportData = dTables

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddUpdateTableInList
' Author    : Adam Waller
' Date      : 4/21/2020
' Purpose   : Updates a table in the collection, adding it if it does not exist.
'---------------------------------------------------------------------------------------
'
Private Sub AddUpdateTableInList(strName As String, lngFormatType As eTableDataExportFormat, blnHidden As Boolean, blnSystem As Boolean, blnOther As Boolean, blnLocal As Boolean)

    Dim rstClone As DAO.Recordset
    Dim rstActive As DAO.Recordset

    Set rstClone = Me.sfrmTableData.Form.RecordsetClone
    Set rstActive = Me.sfrmTableData.Form.Recordset

    With rstActive
        rstClone.FindFirst "TableName='" & Replace$(strName, "'", "''") & "'"
        If rstClone.NoMatch Then
            Me.sfrmTableData.Form.AllowAdditions = True
            .AddNew
            !TableName = strName
            !TableIcon = GetTableIcon(etcOther)
        Else
            .Bookmark = rstClone.Bookmark
            .Edit
        End If
        !FormatType = lngFormatType
        !IsHidden = blnHidden
        !IsSystem = blnSystem
        !IsOther = blnOther
        !IsLocal = blnLocal
        .Update
        Me.sfrmTableData.Form.AllowAdditions = False
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RefreshTableDisplay
' Author    : Adam Waller
' Date      : 4/17/2020
' Purpose   : Update the table data subform display with current filter settings.
'---------------------------------------------------------------------------------------
'
Private Sub RefreshTableDisplay()

    Dim strFilter As String
    Dim strOrderBy As String

    If Me.chkTableShowOther Then
        strOrderBy = strOrderBy & ", IIf([IsOther], 0, 1)"
    Else
        strFilter = strFilter & " AND [IsOther] = 0"
    End If

    If Me.chkTableShowSystem Then
        strOrderBy = strOrderBy & ", IIf([IsSystem], 0, 1)"
    Else
        strFilter = strFilter & " AND [IsSystem] = 0"
    End If

    If Me.chkTableShowHidden Then
        strOrderBy = strOrderBy & ", IIf([IsHidden], 0, 1)"
    Else
        strFilter = strFilter & " AND [IsHidden] = 0"
    End If

    strOrderBy = strOrderBy & ", IIf([IsLocal], 0, 1), [TableName]"

    If Len(strFilter) Then
        strFilter = " WHERE " & Mid$(strFilter, 6)
    End If

    If Len(strOrderBy) Then
        strOrderBy = " ORDER BY " & Mid$(strOrderBy, 3)
    End If

    Me.sfrmTableData.Form.RecordSource = _
        "SELECT d.TableIcon, d.TableName, d.FormatType, d.IsHidden, d.IsSystem, d.IsOther, d.IsLocal " & _
        "FROM tblTableData AS d " & _
        strFilter & _
        strOrderBy

    Me.lblTableShowHidden.Caption = GetCaptionWithCount(T("Show Hidden"), "d.IsHidden = True AND d.IsSystem = " & chkTableShowSystem)
    Me.lblTableShowSystem.Caption = GetCaptionWithCount(T("Show System"), "d.IsSystem = True AND d.IsHidden = " & chkTableShowHidden)
    Me.lblTableShowOther.Caption = GetCaptionWithCount(T("Show Other  "), "d.IsOther  = True")

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetCaptionWithCount
' Author    : bclothier
' Date      : 5/15/2023
' Purpose   : Provides caption with a count appended if non-zero.
'---------------------------------------------------------------------------------------
'
Private Function GetCaptionWithCount(TemplateCaption As String, CountFilter As String) As String
    Dim rs As DAO.Recordset
    Set rs = CodeDb.OpenRecordset( _
        "SELECT COUNT(d.TableName) FROM tblTableData AS d WHERE " & CountFilter, dbOpenSnapshot)
    If rs.EOF = False Then
        If Nz(rs.Fields(0).Value, 0) Then
            GetCaptionWithCount = TemplateCaption & " (" & rs.Fields(0).Value & ")"
            Exit Function
        End If
    End If
    GetCaptionWithCount = TemplateCaption
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTableIcon
' Author    : Adam Waller & Indigo744
' Date      : 11/11/2020
' Purpose   : Get a Unicode icon depending on the type of table.
'---------------------------------------------------------------------------------------
'
Private Function GetTableIcon(ByRef lngColumn As eTableCol) As String
    Select Case lngColumn
        Case etcSystem
            GetTableIcon = ChrW$(11034)
        Case etcLocal
            GetTableIcon = ChrW$(8862)
        Case etcOther
            GetTableIcon = ChrW$(8618)
        Case Else
            GetTableIcon = ChrW$(55356) & ChrW$(57102)
    End Select
End Function
