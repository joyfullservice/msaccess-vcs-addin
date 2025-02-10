Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmOptions
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : NOTE - Options are dynamically loaded to the controls, and then saved
'           : back to the class (and subsequently project) when the user clicks the
'           : save and close button. Most options will not require event coding since
'           : they are mapped to their appropriate option using the control name.
'           : I.e. chkUseFastSave = Options.UseFastSave
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private Enum eTableCol
    etcName = 0
    etcType = 1
    etcHidden = 2
    etcSystem = 3
    etcOther = 4
    etcLocal = 5
    etcLinked = 6
End Enum

Private Enum eMapAction
    emaClassToForm
    emaFormToClass
End Enum


' Dictionary to stash database schemas while managing options.
Public DatabaseSchemas As Dictionary


'---------------------------------------------------------------------------------------
' Procedure : chkTableShowHidden_Click
' Author    : Adam Waller
' Date      : 4/17/2020
' Purpose   : Update the list of tables
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
' Procedure : cmdAddDatabase_Click
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Add an external database connection
'---------------------------------------------------------------------------------------
'
Private Sub cmdAddDatabase_Click()
    DoCmd.OpenForm "frmVCSDatabase"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : chkContributeTranslations_Click
' Author    : Adam Waller
' Date      : 9/2/2024
' Purpose   : Set flag on whether to actively contribute to translation files.
'---------------------------------------------------------------------------------------
'
Private Sub chkContributeTranslations_Click()
    Translation.Contribute = True
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cboLanguage_Click
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Set language
'---------------------------------------------------------------------------------------
'
Private Sub cboLanguage_Click()
    With Translation
        .Language = Nz2(cboLanguage, "en_US")
        .SetLanguage .Language
        MsgBox2 "Language Set", "Changes to language settings will take effect " & _
            "the next time you open this form.", , vbInformation
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdBrowseTranslationsPath_Click
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Set path for translation files
'---------------------------------------------------------------------------------------
'
Private Sub cmdBrowseTranslationsPath_Click()

    Dim strFolder As String

    ' Show a folder picker to select the file with source code.
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .ButtonName = "Select Translations Folder"
        .InitialFileName = StripSlash(txtTranslationsPath) & PathSep
        .Title = "Select Translations Files Folder"
        .Show
        If .SelectedItems.Count > 0 Then
            ' Selected a folder
            strFolder = .SelectedItems(1) & PathSep
            If FSO.FileExists(strFolder & PROJECT_NAME & ".pot") Then
                ' Has translation files
                txtTranslationsPath = strFolder
                Translation.TranslationsPath = strFolder
            Else
                MsgBox2 "Translation files not found", "Required translation files were not found in this folder.", _
                    "You selected: " & .SelectedItems(1), vbExclamation
                DoCmd.Hourglass False
                Exit Sub
            End If
        End If
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdEditDatabase_Click
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Edit an existing database connection
'---------------------------------------------------------------------------------------
'
Private Sub cmdEditDatabase_Click()
    If Len(Nz(lstDatabases)) > 0 Then
        ' Open the form as hidden, then load the properties
        DoCmd.OpenForm "frmVCSDatabase", , , , , acHidden
        With Form_frmVCSDatabase
            .LoadSchema lstDatabases, Me.DatabaseSchemas(Nz(lstDatabases))
            .Visible = True
        End With
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdSyncTranslations_Click
' Author    : Adam Waller
' Date      : 8/23/2024
' Purpose   : Load the translation files (other languages, not the master file),
'           : then export all the translations, including the master file.
'           : Changes to the master file are loaded during a build from source.
'---------------------------------------------------------------------------------------
'
Private Sub cmdSyncTranslations_Click()

    With Translation
        DoCmd.Hourglass True

        ' Only export strings from the installed add-in project. This is where the
        ' majority of the run-time strings will be generated. We don't want to
        ' export the template from the development project because it may not
        ' include these additional strings.
        If CodeProject.FullName = GetAddInFileName Then
            ' Make sure we have strings in the current database before exporting the template.
            If .StringCount > 200 Then .ExportTemplate
        End If

        ' Load template, then other translation files
        .LoadTranslations

        ' Export all translations to include new strings from master
        .ExportTranslations

        DoCmd.Hourglass False
    End With

    MsgBox2 T("Translations Updated"), T("Translation files have been syncronized with the add-in."), , vbInformation

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdTranslations_Click
' Author    : Adam Waller
' Date      : 8/7/2024
' Purpose   : Show the translations tab
'---------------------------------------------------------------------------------------
'
Private Sub cmdTranslations_Click()
    With pgeTranslations
        .Visible = True
        .SetFocus
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Open
' Author    : Adam Waller
' Date      : 10/31/2024
' Purpose   : Check for scenario that may crash Access
'---------------------------------------------------------------------------------------
'
Private Sub Form_Open(Cancel As Integer)
    If Not DatabaseFileOpen Then
        ' The table data subform may crash Access when the options form is closed with no database open.
        MsgBox2 T("No Database Open"), T("Please open a database file before configuring VCS options."), , vbInformation
        Cancel = True
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : lstDatabases_DblClick
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Shortcut to edit the selected database
'---------------------------------------------------------------------------------------
'
Private Sub lstDatabases_DblClick(Cancel As Integer)
    cmdEditDatabase_Click
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdDeleteDatabase_Click
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Delete a database connection
'---------------------------------------------------------------------------------------
'
Private Sub cmdDeleteDatabase_Click()
    Dim strName As String
    strName = Nz(lstDatabases)
    If Len(strName) = 0 Then
        MsgBox2 T("Select a connection to delete"), , , vbExclamation
    Else
        With Me.DatabaseSchemas
            If .Exists(strName) Then
                If MsgBox2(T("Remove Connection?"), T("Are you sure you want to delete '{0}'?", var0:=strName), _
                    T("Click YES to remove or NO to cancel."), vbQuestion + vbYesNo) = vbYes Then
                    .Remove strName
                    RefreshSchemaList
                End If
            End If
        End With
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadSchemas
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Load (or reload) the schemas from a dictionary object
'---------------------------------------------------------------------------------------
'
Public Sub RefreshSchemaList()

    Dim varKey As Variant

    With lstDatabases
        .RowSource = vbNullString
        ' Add header row
        .AddItem T("Name;Description", "lstDatabases.RefreshSchemaList", , "Do not remove the semicolon")
        ' Update list from dictionary
        For Each varKey In Me.DatabaseSchemas.Keys
            .AddItem CStr(varKey) & ";" & Me.DatabaseSchemas(varKey)("Description")
        Next varKey
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdCancel_Click
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Cancel (close) the form.
'---------------------------------------------------------------------------------------
'
Private Sub cmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdClearDefaults_Click
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Clear the default and project settings.
'---------------------------------------------------------------------------------------
'
Private Sub cmdClearDefaults_Click()
    Dim cDefaults As clsOptions
    Set cDefaults = New clsOptions
    cDefaults.SaveOptionsAsDefault
    If MsgBox2(T("Apply to this Project?"), _
        T("The default options have been reset. Would you like these applied to this project as well?"), _
        T("Click 'Yes' to apply the default options, or 'No' to leave the current options."), _
        vbQuestion + vbYesNo, T("Version Control System")) = VbMsgBoxResult.vbYes Then
        cmdRestoreDefaults_Click
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

    ' Reset list of tables
    Set dbs = CodeDb
    dbs.Execute "DELETE FROM tblTableData;", dbFailOnError

    ' Open table to load records
    Set rstTableData = dbs.OpenRecordset("SELECT * FROM tblTableData;", dbOpenDynaset)

    ' Get list of tables if we have a database file open.
    If DatabaseFileOpen Then
        ' ADP does not have local tables. Export data is probably best handled with SQL Server tools.
        If Not (CurrentProject.ProjectType = acADP) Then
            ' Note that Access SQL does not support bitwise "and" operator
            ' (Also known as BAND in ADO) so we will check the bit flags in VBA instead.
            strSql = _
                "SELECT o.Name, o.Type, o.Flags " & _
                "FROM MSysObjects AS o " & _
                "WHERE o.Type IN (1, 4, 6) " & _
                "ORDER BY o.Name;"

            Set rstSource = CurrentDb.OpenRecordset(strSql, dbOpenSnapshot)
            With rstSource
                Do While Not .EOF
                    ' Determine type of table
                    lngFlags = Nz(!Flags, 0)
                    lngType = Nz(!Type, 0)
                    If (lngFlags < 0) Or BitSet(lngFlags, 1) Then
                        ' Don't include read-only or deeply hidden system tables.
                        ' https://isladogs.co.uk/purpose-of-system-tables-2/index.html#TFE
                    Else
                        rstTableData.AddNew
                            rstTableData!TableName = Nz(!Name)
                            rstTableData!Flags = Nz(!Flags)
                            rstTableData!IsSystem = BitSet(lngFlags, 2)
                            rstTableData!IsHidden = BitSet(lngFlags, 8)
                            rstTableData!IsLocal = (lngType = 1)
                            ' Determine table icon
                            rstTableData!TableIcon = GetTableIcon(etcLinked)    ' Default to linked table if no match.
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

    ' Add in the list of saved tables, adding into the sorted location
    If Not Options.TablesToExportData Is Nothing Then
        ' Loop through each table in the saved table list
        For Each varKey In Options.TablesToExportData.Keys
            strName = CStr(varKey)
            strFormat = Options.TablesToExportData.Item(varKey)("Format")
            intFormat = Options.GetTableExportFormat(strFormat)

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

    ' Close recordset after adding records
    rstTableData.Close

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveTableList
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Save the list of tables that should have its data exported to options file.
'---------------------------------------------------------------------------------------
'
Private Sub SaveTableList()

    Dim rstTableData As DAO.Recordset
    Dim dTables As Dictionary
    Dim dTable As Dictionary

    ' Save list of tables to export data
    Set dTables = New Dictionary
    dTables.CompareMode = TextCompare

    ' ADP project type shouldn't export data; set it as an empty list.
    If Not (CurrentProject.ProjectType = acADP) Then
        Set rstTableData = CodeDb.OpenRecordset( _
            "SELECT TableName, FormatType FROM tblTableData " & _
            "WHERE FormatType <> 0 ORDER BY TableName;", dbOpenForwardOnly)
        With rstTableData
            Do Until .EOF
                Set dTable = New Dictionary
                dTable.CompareMode = TextCompare
                dTable("Format") = Options.GetTableExportFormatName(Nz(!FormatType, 0))
                dTables.Add Nz(!TableName), dTable
                .MoveNext
            Loop
            .Close
        End With
    End If

    Set Options.TablesToExportData = dTables
End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddUpdateTableInList
' Author    : Adam Waller
' Date      : 4/21/2020
' Purpose   : Updates the a table in the collection, adding it if it does not exist.
'---------------------------------------------------------------------------------------
'
Private Sub AddUpdateTableInList(strName As String, lngFormatType As eTableDataExportFormat, blnHidden As Boolean, blnSystem As Boolean, blnOther As Boolean, blnLocal As Boolean)

    Dim rstClone As DAO.Recordset
    Dim rstActive As DAO.Recordset

    Set rstClone = Me.sfrmTableData.Form.RecordsetClone
    Set rstActive = Me.sfrmTableData.Form.Recordset

    With rstActive

        ' Look for matching table name
        rstClone.FindFirst "TableName='" & Replace$(strName, "'", "''") & "'"
        If rstClone.NoMatch Then
            ' Add new table to this list
            Me.sfrmTableData.Form.AllowAdditions = True
            .AddNew
            !TableName = strName
            !TableIcon = GetTableIcon(etcOther)
        Else
            .Bookmark = rstClone.Bookmark
            .Edit
        End If

        ' Update remaining fields
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
' Purpose   : Update the subform's display
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

    Dim strSql As String
    strSql = _
        "SELECT d.TableIcon, d.TableName, d.FormatType, d.IsHidden, d.IsSystem, d.IsOther, d.IsLocal " & _
        "FROM tblTableData AS d " & _
        strFilter & _
        strOrderBy

    Me.sfrmTableData.Form.RecordSource = strSql

    ' Update captions with counts
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
' Procedure : cmdOpenInstallFolder_Click
' Author    : Adam Waller
' Date      : 7/6/2023
' Purpose   : Open the installation folder
'---------------------------------------------------------------------------------------
'
Private Sub cmdOpenInstallFolder_Click()
    Application.FollowHyperlink modInstall.GetInstallSettings.strInstallFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdRestoreDefaults_Click
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Restore the default options to this project.
'---------------------------------------------------------------------------------------
'
Private Sub cmdRestoreDefaults_Click()
    Options.LoadDefaultOptions
    MapControlsToOptions emaClassToForm
    MsgBox2 T("Default Options Restored"), _
        T("The system default options have been restored to this project."), _
        T("Be sure to click Save and Close when you are finished making changes."), _
        vbInformation, T("Version Control System")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdSaveAndClose_Click
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Save and close the options dialog
'---------------------------------------------------------------------------------------
'
Private Sub cmdSaveAndClose_Click()

    ' Make sure we actually have a file open
    If Not DatabaseFileOpen Then
        MsgBox2 T("No Database File Open"), _
            T("You must have a database file open to save VCS options to a source folder."), _
            T("Please open a database file before saving options for a project."), vbExclamation
        Exit Sub
    End If

    ' Save options and close.
    MapControlsToOptions emaFormToClass
    Options.SaveOptionsForProject
    DoCmd.Close acForm, Me.Name

    ' Update main form if options changed.
    If IsLoaded(acForm, "frmVCSMain", True) Then Form_frmVCSMain.Form_Load

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdSaveAsDefault_Click
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Save current options as default for projects.
'---------------------------------------------------------------------------------------
'
Private Sub cmdSaveAsDefault_Click()

    Dim strPath As String

    ' Note that we can't save an absolute path as default, or we will potentially
    ' create some major issues with source files being overwritten and lost.
    strPath = Nz(txtExportFolder)
    If strPath <> vbNullString Then
        If InStr(1, strPath, "%dbName%", vbTextCompare) < 1 Then
            MsgBox2 T("Invalid Export Path for Default"), _
                T("If you specify an absolute or relative Export Path as a default option," & vbCrLf & _
                "you must include the %dbName% placeholder to keep the paths unique."), _
                T("Please update the Export Path and try again."), vbExclamation
            Exit Sub
        End If
    End If

    ' Load the options from the form and save as default
    MapControlsToOptions emaFormToClass
    Options.SaveOptionsAsDefault

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Load options for this project
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    Dim intFormat As eTableDataExportFormat
    Dim intSanitizeLevel As eSanitizeLevel

    ' Set version
    lblVersion.Caption = GetVCSVersion

    ' Apply translation
    cboLanguage = Translation.Language
    Translation.ApplyTo Me

    MapControlsToOptions emaClassToForm
    RefreshTableDisplay
    RefreshSchemaList

    ' Load list of table data export formats
    Dim frmTableData As Form_frmVCSTableData
    Set frmTableData = Me.sfrmTableData.Form
    With frmTableData.cboFormatType
        .RowSource = vbNullString
        For intFormat = eTableDataExportFormat.etdNoData To eTableDataExportFormat.[_Last] - 1
            .AddItem intFormat & ";""" & Options.GetTableExportFormatName(intFormat) & """", intFormat
        Next intFormat
        Me.cboFormatTypeForOther.RowSource = .RowSource
        Me.cboFormatTypeForOther.RemoveItem etdNoData
    End With

    ' Load general sanitize options
    With Me.cboSanitizeLevel
        .RowSource = vbNullString
        For intSanitizeLevel = eslMinimal To (eSanitizeLevel.[_Last] - 1)
            .AddItem intSanitizeLevel & ";" & Options.GetSanitizeLevelName(intSanitizeLevel)
        Next intSanitizeLevel
    End With

    ' Load color sanitize options
    With Me.cboSanitizeColors
        .RowSource = vbNullString
        For intSanitizeLevel = eslNone To (eSanitizeLevel.[_Last] - 1)
            .AddItem intSanitizeLevel & ";" & Options.GetSanitizeLevelName(intSanitizeLevel)
        Next intSanitizeLevel
    End With

    ' Load language options
    With Translation
        cboLanguage = .Language
        chkContributeTranslations = .Contribute
        txtTranslationsPath = .TranslationsPath
    End With

    ' Make form resizable (helpful with table selection when many tables are listed)
    MakeDialogResizable Me

    ' Set inital column sizing for table data
    DoEvents
    Form_frmVCSTableData.Form_Resize

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MapControlsToOptions
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Map the form controls to the options, performing the specified action.
'---------------------------------------------------------------------------------------
'
Private Sub MapControlsToOptions(eAction As eMapAction)

    Dim pge As Access.Page
    Dim ctl As Access.Control
    Dim strKey As String
    Dim dSettings As Dictionary

    ' Loop through each page
    For Each pge In tabOptions.Pages
        For Each ctl In pge.Controls
            Select Case ctl.ControlType
                Case Access.acCheckBox, Access.acTextBox, Access.acComboBox
                    strKey = Mid$(ctl.Name, 4)
                    If pge Is pgePrinterSettings Then
                        ' Export print options
                        If eAction = emaClassToForm Then
                            ctl = Options.ExportPrintSettings(strKey)
                        ElseIf eAction = emaFormToClass Then
                            Set dSettings = Options.ExportPrintSettings
                            dSettings(strKey) = CBool(ctl)
                        End If
                    Else
                        Select Case ctl.Name
                            Case Me.chkTableShowHidden.Name, Me.chkTableShowSystem.Name, Me.chkTableShowOther.Name, _
                                Me.cboFormatTypeForOther.Name, Me.txtOtherTableName.Name, _
                                Me.cboLanguage.Name, Me.chkContributeTranslations.Name, Me.txtTranslationsPath.Name
                                ' Skip these exceptions.
                            Case Me.cboDiffTool.Name
                                ' This is saved on the computer profile (registry)
                                If eAction = emaClassToForm Then
                                    ctl = Diff.ToolName
                                Else
                                    Diff.ToolName = Nz(ctl)
                                End If
                            Case Me.cboOpenRepository.Name
                                ' Also saved to registry
                                If eAction = emaClassToForm Then
                                    ctl = Nz2(GetSetting(PROJECT_NAME, "Options", "Open Repository Tool"), 0)
                                Else
                                    SaveSetting PROJECT_NAME, "Options", "Open Repository Tool", CStr(Nz(cboOpenRepository, 0))
                                End If
                            Case Else
                                ' Get option name from control name following prefix.
                                If eAction = emaClassToForm Then
                                    ctl = CallByName(Options, strKey, VbGet)
                                ElseIf eAction = emaFormToClass Then
                                    ' Check for any hooks on option change
                                    OnOptionChange strKey, Nz(ctl.Value)
                                    ' Set the option value
                                    CallByName Options, strKey, VbLet, Nz(ctl.Value)
                                End If
                        End Select
                    End If
                Case Access.acListBox
                    ' Special handling for lists
                Case Else
                    ' Ignore other controls
            End Select
        Next ctl
    Next

    ' Table list for selecting tables to save data.
    If eAction = emaClassToForm Then
        LoadTableList
    ElseIf eAction = emaFormToClass Then
        SaveTableList
    End If

    ' Database schemas
    If eAction = emaClassToForm Then
        Set DatabaseSchemas = CloneDictionary(Options.SchemaExports)
    ElseIf eAction = emaFormToClass Then
        Set Options.SchemaExports = CloneDictionary(DatabaseSchemas)
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : OnOptionChange
' Author    : Adam Waller
' Date      : 11/9/2023
' Purpose   : A hook to run special code or processing when specific options are changed
'           : from their existing values. Add any specific rules here.
'---------------------------------------------------------------------------------------
'
Private Sub OnOptionChange(strName As String, varNewValue As Variant)

    Dim blnChanged As Boolean

    ' Determine if the option was changed
    blnChanged = Not (CVar(CallByName(Options, strName, VbGet)) = varNewValue)
    If Not blnChanged Then Exit Sub

    ' Define actual rules here
    Select Case strName

        ' If a user turns on the option to split files
        Case "SplitLayoutFromVBA"
            If varNewValue = True Then
                If Git.Installed Then
                    If Git.IsInsideRepository Then
                        ' Prompt user with suggestion
                        If MsgBox2(T("May I make a Suggestion?"), _
                            T("This project appears to be within a Git repository. This add-in includes a special utility " & _
                            "that can split the files (layout and VBA) while preserving this history of previous changes in BOTH files."), _
                            T("Would you like to see additional information on this from the wiki?"), vbQuestion + vbYesNo) = vbYes Then
                            FollowHyperlink "https://github.com/joyfullservice/msaccess-vcs-addin/wiki/Split-Files"
                        End If
                    End If
                End If
            End If

    End Select

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdAddOtherTableData_Click
' Author    : Adam Waller
' Date      : 4/16/2020
' Purpose   : Select another table not used in the current database
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
' Procedure : cmdPrintSettingsOptions_Click
' Author    : Adam Waller
' Date      : 11/9/2020
' Purpose   : View advanced options for saving print settings.
'---------------------------------------------------------------------------------------
'
Private Sub cmdPrintSettingsOptions_Click()
    With pgePrinterSettings
        .Visible = True
        .SetFocus
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdUninstall_Click
' Author    : Adam Kauffman
' Date      : 5/25/2020
' Purpose   : Remove the add-in.
'---------------------------------------------------------------------------------------
'
Private Sub cmdUninstall_Click()
    UninstallVCSAddin
End Sub


'---------------------------------------------------------------------------------------
' Procedure : txtExportFolder_BeforeUpdate
' Author    : Adam Waller
' Date      : 5/6/2021
' Purpose   : Make sure we have a valid entry, blank, absolute path, or relative path.
'---------------------------------------------------------------------------------------
'
Private Sub txtExportFolder_BeforeUpdate(Cancel As Integer)

    Dim strPath As String

    strPath = Nz(txtExportFolder)
    If strPath <> vbNullString Then
        If (Left(strPath, 1) = PathSep) Or _
            (InStr(2, strPath, ":" & PathSep) > 0) Then
            ' Looks like a valid path
        Else
            MsgBox2 T("Invalid Export Folder"), _
                T("This does not appear to be a valid relative or absolute path."), _
                T("Please see the wiki documentation for more detail and examples."), vbExclamation
            Cancel = True
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetTableIcon
' Author    : Adam Waller & Indigo744
' Date      : 11/11/2020
' Purpose   : Get an icon depending on the type of table
'---------------------------------------------------------------------------------------
'
Private Function GetTableIcon(ByRef lngColumn As eTableCol) As String
    Select Case lngColumn
        Case etcSystem
            ' System tables
            ' Uses symbol DOTTED SQUARE
            ' http://www.fileformat.info/info/unicode/char/2b1a/index.htm
            GetTableIcon = ChrW$(11034)
        Case etcLocal
            ' Local tables
            ' Uses symbol SQUARED PLUS
            ' http://www.fileformat.info/info/unicode/char/229e/index.htm
            GetTableIcon = ChrW$(8862)
        Case etcOther
            ' "Other" tables
            ' Uses symbol RIGHTWARDS ARROW WITH HOOK
            ' http://www.fileformat.info/info/unicode/char/21aa/index.htm
            GetTableIcon = ChrW$(8618)
        Case Else
            ' Anything else would be a linked table
            ' Uses symbol EARTH GLOBE AMERICAS
            ' https://www.fileformat.info/info/unicode/char/1f30e/index.htm
            GetTableIcon = ChrW$(55356) & ChrW$(57102)
    End Select
End Function
