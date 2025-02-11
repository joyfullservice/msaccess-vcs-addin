Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' Store original name, just in case we rename an existing entry.
Private m_strOriginalName As String


'---------------------------------------------------------------------------------------
' Procedure : LoadSchema
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Load a schema by name
'---------------------------------------------------------------------------------------
'
Public Sub LoadSchema(strName As String, dSchema As Dictionary)

    Dim dParams As Dictionary

    ' Merge values into defaults to ensure that all values are defined
    Set dParams = GetDefaults
    MergeDictionary dParams, dSchema
    MergeDotEnv strName, dParams

    ' Save original name, just in case it is renamed
    m_strOriginalName = strName

    ' Load values from dictionary parameters
    txtName = strName
    chkEnabled = dParams("Enabled")
    cboType = dParams("DatabaseType")
    chkUtcDates = dParams("UtcDateTime")
    txtDescription = dParams("Description")
    txtFilter = ParseFilter(dParams("Filter"))
    cboConnect = dParams("Connect")
    chkSaveDotEnv = dParams("UseDotEnv")
    chkAttemptADOConvert = dParams("AttemptConvert")

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ParseFilter
' Author    : Adam Waller
' Date      : 1/9/2024
' Purpose   : Parse the filter
'---------------------------------------------------------------------------------------
'
Private Function ParseFilter(varValue As Variant) As String

    Dim varItem As Variant

    ' Convert collection to filter lines
    With New clsConcat
        .AppendOnAdd = vbCrLf
        For Each varItem In varValue
            .Add CStr(varItem)
        Next varItem
        If .Length > 2 Then .Remove 2
        ParseFilter = .GetStr
    End With

End Function


'---------------------------------------------------------------------------------------
' Procedure : MergeDotEnv
' Author    : Adam Waller
' Date      : 7/31/2023
' Purpose   : Merge any specified .env params into dictionary. (Will override schema
'           : options saved in vcs-options.json)
'---------------------------------------------------------------------------------------
'
Private Sub MergeDotEnv(strSchemaName As String, ByRef dParams As Dictionary)
    Dim strFile As String
    strFile = BuildPath2(Options.GetExportFolder & "databases", GetSafeFileName(strSchemaName), ".env")
    If FSO.FileExists(strFile) Then
        With New clsDotEnv
            .LoadFromFile strFile
            .MergeIntoDictionary dParams, False
        End With
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveConnectionString
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Saves the connection string to a .env file.
'---------------------------------------------------------------------------------------
'
Private Sub SaveConnectionStringToFile()

    Dim strFile As String

    ' Guard clause safety check
    If Nz(txtName) = vbNullString Or Nz(cboConnect) = vbNullString Then Exit Sub

    ' Update the value in the .env file. (Creating the file, if needed.)
    strFile = BuildPath2(Options.GetExportFolder & "databases", GetSafeFileName(Nz(txtName)), ".env")
    With New clsDotEnv
        ' Reload file so we preserve existing values
        .LoadFromFile strFile
        .SetVar "CONNECT", Nz(cboConnect)
        .SaveToFile strFile
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cboConnect_Click
' Author    : Adam Waller
' Date      : 2/11/2025
' Purpose   : Encourage the use of a .env file if passwords are identified in the
'           : connection string.
'---------------------------------------------------------------------------------------
'
Private Sub cboConnect_Click()
    Dim strTest As String
    strTest = Nz(cboConnect)
    If Len(strTest) Then
        If strTest Like "*password=*" Or strTest Like "*pwd=*" Then
            ' Appears to have a saved password
            If Not Me.chkSaveDotEnv Then
                If MsgBox2("Use .env File for Connection String?", _
                    "Using a .env file to store sensitive credentials is highly recommended " & _
                    "when your connection string contains a password.", _
                    "This .env file should be excluded from version control.", vbQuestion + vbYesNo) = vbYes Then
                    Me.chkSaveDotEnv = True
                End If
            End If
        End If
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdCancel_Click
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Close this form
'---------------------------------------------------------------------------------------
'
Private Sub cmdCancel_Click()
    DoCmd.Close acForm, Me.Name
End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdExamples_Click
' Author    : Adam Waller
' Date      : 8/2/2023
' Purpose   : Show syntax examples on the Wiki
'---------------------------------------------------------------------------------------
'
Private Sub cmdExamples_Click()

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdSaveAndClose_Click
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Save the schema back to the options form.
'---------------------------------------------------------------------------------------
'
Private Sub cmdSaveAndClose_Click()
    If SaveConnection Then
        Form_frmVCSOptions.RefreshSchemaList
        DoCmd.Close acForm, Me.Name
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveConnection
' Author    : Adam Waller
' Date      : 7/21/2023
' Purpose   : Save the current connection, return true if successful.
'---------------------------------------------------------------------------------------
'
Private Function SaveConnection() As Boolean

    Dim dSchema As Dictionary
    Dim strKey As String

    If Not PassedValidation Then Exit Function

    If IsLoaded(acForm, "frmVCSOptions") Then
        With Form_frmVCSOptions

            ' Make sure we have a dictionary object
            If .DatabaseSchemas Is Nothing Then Set .DatabaseSchemas = New Dictionary

            ' Save to options form
            With .DatabaseSchemas

                ' Get a reference to dictionary object
                strKey = Nz(txtName)
                If Not .Exists(strKey) Then
                    ' Could be a rename
                    Set dSchema = New Dictionary
                    .Add strKey, dSchema
                    ' Remove any previous entry
                    If Len(m_strOriginalName) Then
                        If .Exists(m_strOriginalName) Then .Remove m_strOriginalName
                    End If
                End If

                ' Load form values
                SetParamsFromForm .Item(strKey)

                ' Connection string
                If chkSaveDotEnv Then
                    CheckGitignoreDotEnv
                    ' Save connection string to .env file
                    SaveConnectionStringToFile
                    ' Remove connect parameter from dictionary
                    If .Item(strKey).Exists("Connect") Then .Item(strKey).Remove "Connect"
                End If
            End With
        End With

        ' Return success
        SaveConnection = True
    Else
        MsgBox2 "Options form not found", "The Options form must be open to save changes to external database connections", , vbExclamation
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : SetParamsFromForm
' Author    : Adam Waller
' Date      : 7/31/2023
' Purpose   : Sets the dictionary parameters based on the current form values.
'           : (This is especially helpful for testing filters before saving an entry.)
'---------------------------------------------------------------------------------------
'
Private Sub SetParamsFromForm(ByRef dParams As Dictionary)
    With dParams
        .Item("Enabled") = CBool(chkEnabled)
        .Item("DatabaseType") = CInt(cboType)
        .Item("Description") = Nz(txtDescription)
        .Item("UtcDateTime") = CBool(chkUtcDates)
        .Item("Connect") = Nz(cboConnect)
        .Item("UseDotEnv") = CBool(chkSaveDotEnv)
        .Item("AttemptConvert") = CBool(chkAttemptADOConvert)
        Set .Item("Filter") = ToCollection(Nz(txtFilter))
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ToCollection
' Author    : Adam Waller
' Date      : 1/9/2024
' Purpose   : Convert a string list of items to a collection. (Split on vbCrLf)
'---------------------------------------------------------------------------------------
'
Private Function ToCollection(strFilter) As Collection

    Dim colRules As Collection
    Dim varLine As Variant

    Set colRules = New Collection
    For Each varLine In Split(strFilter, vbCrLf)
        colRules.Add CStr(varLine)
    Next varLine
    Set ToCollection = colRules

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDefaults
' Author    : Adam Waller
' Date      : 7/31/2023
' Purpose   : Set the default values for connection parameters
'---------------------------------------------------------------------------------------
'
Private Function GetDefaults() As Dictionary
    Set GetDefaults = New Dictionary
    With GetDefaults
        .CompareMode = TextCompare
        .Item("Enabled") = True
        .Item("DatabaseType") = 1
        .Item("Description") = vbNullString
        .Item("UtcDateTime") = False
        .Item("Connect") = vbNullString
        .Item("UseDotEnv") = True
        .Item("AttemptConvert") = True
        Set .Item("Filter") = New Collection
    End With
End Function


'---------------------------------------------------------------------------------------
' Procedure : PassedValidation
' Author    : Adam Waller
' Date      : 7/20/2023
' Purpose   : Return true if we pass validation on the form to save the entry.
'---------------------------------------------------------------------------------------
'
Private Function PassedValidation() As Boolean

    Dim strMsg As String

    ' TODO: Could add more validation for filter entries

    If Len(Nz(cboConnect)) < 5 Then strMsg = "Please select or enter a connection string for database"
    If Nz(cboType, -1) < 0 Then strMsg = "Please select database type"
    If Len(Nz(txtName)) = 0 Then strMsg = "Connection name is required"

    If Len(strMsg) Then
        MsgBox2 "Please fix validation issues to continue", strMsg, "See online wiki for additional documentation", vbExclamation
    Else
        PassedValidation = True
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : cmdTest_Click
' Author    : Adam Waller
' Date      : 7/21/2023
' Purpose   : Test the current filter and return the number of objects found.
'---------------------------------------------------------------------------------------
'
Private Sub cmdTest_Click()

    Dim cSchema As IDbSchema
    Dim lngCount As Long
    Dim dblStart As Double
    Dim dParams As Dictionary

    txtFilter.SetFocus
    txtFilter.SelStart = 9999
    If Not PassedValidation Then Exit Sub

    Select Case cboType
        Case eDatabaseServerType.estMsSql
            Set cSchema = New clsSchemaMsSql
        Case eDatabaseServerType.estMySql
            Set cSchema = New clsSchemaMySql
    End Select

    ' Retrieve object count from server.
    If Not cSchema Is Nothing Then
        Set dParams = New Dictionary
        dParams.CompareMode = TextCompare
        SetParamsFromForm dParams
        cSchema.Initialize dParams
        dblStart = Perf.MicroTimer
        DoCmd.Hourglass True
        lngCount = cSchema.ObjectCount(False)
        DoCmd.Hourglass False
        MsgBox2 lngCount & " Objects Found", "A total of " & lngCount & " database objects were retrieved in " & _
            Round(Perf.MicroTimer - dblStart, 2) & " seconds.", , vbInformation
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CheckGitignoreDotEnv
' Author    : Adam Waller
' Date      : 7/31/2023
' Purpose   : If the project appears to be a .git repository, check to see if .env
'           : appears in the .gitignore file.
'           : (This is not a comprehensive test, but just an extra aid for most common
'           :  scenarios to help users avoid inadvertently comitting a .env file to
'           :  their version control system.)
'---------------------------------------------------------------------------------------
'
Private Sub CheckGitignoreDotEnv()

    Dim strPath As String
    Dim strContent As String

    ' Guess at the standard location for a .gitignore file
    strPath = Options.GetExportFolder & "..\.gitignore"
    If FSO.FileExists(strPath) Then
        strContent = ReadFile(strPath)
        If Len(strContent) Then
            If InStr(1, strContent, ".env", vbTextCompare) = 0 Then
                MsgBox2 "Potentially Sensitive File", _
                    "Please note: .env files should not be committed to version control.", _
                    "To avoid exposing credentials to your repository, please exclude .env files in .gitignore", _
                    vbExclamation
            End If
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 8/1/2023
' Purpose   : Load in sample connection strings from the current database.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    Translation.ApplyTo Me
    LoadSampleConnectionStrings
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadSampleConnectionStrings
' Author    : Adam Waller
' Date      : 8/1/2023
' Purpose   : Loads in some sample connection strings from the current database (from
'           : tables or pass-through queries) that can be selected by the user.
'---------------------------------------------------------------------------------------
'
Private Sub LoadSampleConnectionStrings()

    Dim dItems As Dictionary
    Dim varKey As Variant
    Dim varConn As Variant
    Dim strCurrent As String

    ' Take advantage of our connection class to retrieve some example
    ' connection strings from the current database.
    With New clsDbConnection
        Set dItems = .GetDictionary
    End With

    ' Save current value, and rebuild list
    strCurrent = Nz(cboConnect)
    cboConnect.RowSource = vbNullString

    ' Loop through the connection strings
    If Not dItems Is Nothing Then
        For Each varKey In dItems.Keys
            For Each varConn In dItems(varKey)
                If Len(varConn) > 10 Then
                    ' Looks like a connection string. Add to list.
                    cboConnect.AddItem """" & varConn & """"
                End If
            Next varConn
        Next varKey
    End If

    ' Restore original value
    cboConnect = strCurrent

End Sub
