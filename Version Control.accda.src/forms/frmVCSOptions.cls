Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSOptions
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Main options form with left navigation and subform-per-section layout.
'           : Holds a working copy of the options class (m_Options) that subforms
'           : read from and write to via the OptionsInstance property. Changes are
'           : committed to file only when the user clicks Save & Close.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


' Instance of database schemas dictionary to maintain modified schemas
Public DatabaseSchemas As Dictionary

Private m_Options As clsOptions

' Pending registry settings (deferred until Save & Close)
Private m_strDiffTool As String
Private m_lngOpenRepository As Long


'---------------------------------------------------------------------------------------
' Procedure : OptionsInstance
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Public property for subforms to access the working options instance.
'---------------------------------------------------------------------------------------
'
Public Property Get OptionsInstance() As clsOptions
    Set OptionsInstance = m_Options
End Property


'---------------------------------------------------------------------------------------
' Procedure : DiffTool
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Pending diff tool setting (registry, deferred save).
'---------------------------------------------------------------------------------------
'
Public Property Get DiffTool() As String
    DiffTool = m_strDiffTool
End Property

Public Property Let DiffTool(strValue As String)
    m_strDiffTool = strValue
End Property


'---------------------------------------------------------------------------------------
' Procedure : OpenRepository
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Pending open repository tool setting (registry, deferred save).
'---------------------------------------------------------------------------------------
'
Public Property Get OpenRepository() As Long
    OpenRepository = m_lngOpenRepository
End Property

Public Property Let OpenRepository(lngValue As Long)
    m_lngOpenRepository = lngValue
End Property


'---------------------------------------------------------------------------------------
' Procedure : Form_Open
' Author    : Adam Waller
' Date      : 10/31/2024
' Purpose   : Check for scenario that may crash Access
'---------------------------------------------------------------------------------------
'
Private Sub Form_Open(Cancel As Integer)
    If Not DatabaseFileOpen Then
        MsgBox2 T("No Database Open"), T("Please open a database file before configuring VCS options."), , vbInformation
        Cancel = True
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Initialize the working options copy and load the first section.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()

    ' Set version label
    lblVersion.Caption = GetVCSVersion

    ' Create working copy of options (loads defaults + saved defaults + project options)
    Set m_Options = New clsOptions
    m_Options.LoadProjectOptions

    ' Load current registry settings as pending values
    m_strDiffTool = Diff.ToolName
    m_lngOpenRepository = CLng(Nz2(GetSetting(PROJECT_NAME, "Options", "Open Repository Tool"), 0))

    ' Apply translation
    Translation.ApplyTo Me

    ' Make form resizable
    MakeDialogResizable Me

    ' Navigate to the first section (select the first toggle button)
    Me.fraNav = Me.tglGeneral.OptionValue
    fraNav_AfterUpdate

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Close
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Fires when the form is closed.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Close()
    Set m_Options = Nothing
    Operation.Finish
End Sub


'---------------------------------------------------------------------------------------
' Procedure : fraNav_AfterUpdate
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Handle navigation between option sections. Saves the current subform
'           : state before switching to the newly selected section.
'---------------------------------------------------------------------------------------
'
Private Sub fraNav_AfterUpdate()

    ' Save current subform state before switching
    SaveCurrentSubform

    ' Load the selected section form (toggle name tglXxx maps to frmVCSOptionsXxx)
    Me.subOptionsDetail.SourceObject = "Form.frmVCSOptions" & GetSelectedNavName

    ' Set focus to subform (first element)
    Me.subOptionsDetail.SetFocus

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetSelectedNavName
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Return the section name from the currently selected toggle button.
'           : Derives the name by stripping the "tgl" prefix from the toggle button
'           : whose OptionValue matches the option group's current value.
'---------------------------------------------------------------------------------------
'
Private Function GetSelectedNavName() As String
    Dim ctl As Access.Control
    For Each ctl In Me.fraNav.Controls
        If ctl.ControlType = acToggleButton Then
            If ctl.OptionValue = Me.fraNav.Value Then
                GetSelectedNavName = Mid$(ctl.Name, 4)
                Exit Function
            End If
        End If
    Next ctl
End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveCurrentSubform
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Save the current subform's control values back to the working options
'           : instance via the IOptionsSection interface.
'---------------------------------------------------------------------------------------
'
Public Sub SaveCurrentSubform()
    If Len(Nz(Me.subOptionsDetail.SourceObject)) > 0 Then
        Dim sec As IOptionsSection
        Set sec = Me.subOptionsDetail.Form
        sec.SaveOptions
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RefreshDatabaseSchemaList
' Author    : Adam Waller
' Date      : 7/9/2026
' Purpose   : Refresh the External Databases list on the loaded subform instance.
'           : frmVCSDatabase must not call Form_frmVCSOptionsDatabases directly — that
'           : default instance is not the subform embedded in subOptionsDetail.
'---------------------------------------------------------------------------------------
'
Public Sub RefreshDatabaseSchemaList()
    If InStr(1, Nz(Me.subOptionsDetail.SourceObject), "frmVCSOptionsDatabases", vbTextCompare) > 0 Then
        Me.subOptionsDetail.Form.RefreshSchemaList
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : NavigateTo
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Public method for subforms to request navigation to a different section.
'           : Finds the toggle button matching the section name and selects it.
'           : (e.g. "Translations..." link on General navigating to Translation)
'           : Optional strFocusControl sets focus on a control in the loaded subform.
'---------------------------------------------------------------------------------------
'
Public Sub NavigateTo(strSection As String, Optional strFocusControl As String)
    Dim ctl As Access.Control
    For Each ctl In Me.fraNav.Controls
        If ctl.ControlType = acToggleButton Then
            If Mid$(ctl.Name, 4) = strSection Then
                Me.fraNav = ctl.OptionValue
                fraNav_AfterUpdate
                If Len(strFocusControl) > 0 Then
                    Me.subOptionsDetail.Form.Controls(strFocusControl).SetFocus
                End If
                Exit Sub
            End If
        End If
    Next ctl
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdSaveAndClose_Click
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Save working options to file and close.
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

    ' Save current subform state to working options
    SaveCurrentSubform

    ' Commit working options to the global options and save to file
    Set Options = m_Options
    Options.SaveOptionsForProject

    ' Commit pending registry settings
    Diff.ToolName = m_strDiffTool
    SaveSetting PROJECT_NAME, "Options", "Open Repository Tool", CStr(m_lngOpenRepository)

    DoCmd.Close acForm, Me.Name

    ' Update main form if options changed.
    If IsLoaded(acForm, "frmVCSMain", True) Then Form_frmVCSMain.Form_Load

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdCancel_Click
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Cancel (close) the form, discarding the working options copy.
'---------------------------------------------------------------------------------------
'
Private Sub cmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub
