Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : Form_frmVCSOptionsDatabases
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : External Databases options section: manage connections to external
'           : database schemas for version control tracking.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Implements IOptionsSection

' Working copy of database schemas for this options session.
Private m_dSchemas As Dictionary


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Load schemas from the working options instance.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    LoadOptions
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
' Purpose   : Load database schemas from the working options instance.
'---------------------------------------------------------------------------------------
'
Private Sub LoadOptions()
    ' Clone from working options into the main form's public dictionary, then point
    ' m_dSchemas at the same object. This allows frmVCSDatabase.SaveConnection to
    ' write directly to the dictionary that RefreshSchemaList reads from.
    Set Form_frmVCSOptions.DatabaseSchemas = CloneDictionary(Form_frmVCSOptions.OptionsInstance.SchemaExports)
    Set m_dSchemas = Form_frmVCSOptions.DatabaseSchemas
    RefreshSchemaList
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptions
' Author    : Adam Waller
' Date      : 3/17/2026
' Purpose   : Save database schemas back to the working options instance.
'---------------------------------------------------------------------------------------
'
Private Sub SaveOptions()
    Set Form_frmVCSOptions.OptionsInstance.SchemaExports = CloneDictionary(m_dSchemas)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RefreshSchemaList
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Refresh the list box display from the schemas dictionary.
'---------------------------------------------------------------------------------------
'
Public Sub RefreshSchemaList()

    Dim varKey As Variant

    With Me.lstDatabases
        .RowSource = vbNullString
        .AddItem T("Name;Description", "lstDatabases.RefreshSchemaList", , "Do not remove the semicolon")
        For Each varKey In m_dSchemas.Keys
            .AddItem CStr(varKey) & ";" & m_dSchemas(varKey)("Description")
        Next varKey
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdAddDatabase_Click
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Add an external database connection.
'---------------------------------------------------------------------------------------
'
Private Sub cmdAddDatabase_Click()
    DoCmd.OpenForm "frmVCSDatabase"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdEditDatabase_Click
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Edit the selected database connection.
'---------------------------------------------------------------------------------------
'
Private Sub cmdEditDatabase_Click()
    If Len(Nz(lstDatabases)) > 0 Then
        DoCmd.OpenForm "frmVCSDatabase", , , , , acHidden
        With Form_frmVCSDatabase
            .LoadSchema lstDatabases, m_dSchemas(Nz(lstDatabases))
            .Visible = True
        End With
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : lstDatabases_DblClick
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Shortcut to edit the selected database.
'---------------------------------------------------------------------------------------
'
Private Sub lstDatabases_DblClick(Cancel As Integer)
    cmdEditDatabase_Click
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdDeleteDatabase_Click
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Delete the selected database connection.
'---------------------------------------------------------------------------------------
'
Private Sub cmdDeleteDatabase_Click()
    Dim strName As String
    strName = Nz(lstDatabases)
    If Len(strName) = 0 Then
        MsgBox2 T("Select a connection to delete"), , , vbExclamation
    Else
        If m_dSchemas.Exists(strName) Then
            If MsgBox2(T("Remove Connection?"), T("Are you sure you want to delete '{0}'?", var0:=strName), _
                T("Click YES to remove or NO to cancel."), vbQuestion + vbYesNo) = vbYes Then
                m_dSchemas.Remove strName
                RefreshSchemaList
            End If
        End If
    End If
End Sub
