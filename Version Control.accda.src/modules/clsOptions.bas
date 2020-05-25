Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Note, this class could be described as Options, Settings or Preferences. After some
' deliberation and review, I chose to go with Options because it is most widely used
' in Microsoft products (Office, Visual Studio, VBA IDE, etc...)
' Further reading: https://english.stackexchange.com/questions/59058/

Private Const cstrOptionsFilename As String = "vcs-options.json"

' Options
Public ExportFolder As String
Public ShowDebug As Boolean
Public UseFastSave As Boolean
Public SavePrintVars As Boolean
Public SaveQuerySQL As Boolean
Public SaveTableSQL As Boolean
Public StripPublishOption As Boolean
Public AggressiveSanitize As Boolean
Public TablesToExportData As Scripting.Dictionary
Public RunBeforeExport As String
Public RunAfterExport As String
Public RunAfterBuild As String
Public KeyName As String

Private m_colOptions As New Collection


'---------------------------------------------------------------------------------------
' Procedure : LoadDefaults
' Author    : Adam Waller
' Date      : 2/12/2020
' Purpose   : Loads the default values. Define system defaults here.
'           : (Some values not defined here when they initialize to the default state.)
'---------------------------------------------------------------------------------------
'
Public Sub LoadDefaults()

    With Me
        .ExportFolder = vbNullString
        .ShowDebug = False
        .UseFastSave = True
        .SavePrintVars = True
        .SaveQuerySQL = True
        .SaveTableSQL = True
        .StripPublishOption = True
        .AggressiveSanitize = True
        .KeyName = "MSAccessVCS"
        Set .TablesToExportData = New Scripting.Dictionary
        ' Save specific tables by default
        AddTableToExportData "USysRibbons", etdTabDelimited
        AddTableToExportData "USysRegInfo", etdTabDelimited
    End With
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddTableToExportData
' Author    : Adam Waller
' Date      : 4/17/2020
' Purpose   : Add a table to the list of saved tables
'---------------------------------------------------------------------------------------
'
Public Sub AddTableToExportData(strName As String, intExportFormat As eTableDataExportFormat)
    
    Dim strFormat(etdTabDelimited To etdXML)
    Dim dTable As Scripting.Dictionary
    
    Set dTable = New Scripting.Dictionary
    
    strFormat(etdTabDelimited) = "TabDelimited"
    strFormat(etdXML) = "XMLFormat"
    With Me.TablesToExportData
        Set .Item(strName) = dTable
        .Item(strName)("Format") = GetTableExportFormatName(intExportFormat)
        ' Could add ExcludeColumns here later...
    End With
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptionsToFile
' Author    : Adam Waller
' Date      : 2/12/2020
' Purpose   : Save the loaded Options to a file in JSON format
'---------------------------------------------------------------------------------------
'
Public Sub SaveOptionsToFile(strPath As String)
    WriteFile modJsonConverter.ConvertToJson(SerializeOptions, JSON_WHITESPACE) & vbCrLf, strPath
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptionsAsDefault
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Save these options as default for projects.
'---------------------------------------------------------------------------------------
'
Public Sub SaveOptionsAsDefault()
    Me.SaveOptionsToFile CodeProject.Path & "\" & FSO.GetBaseName(CodeProject.Name) & ".json"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptionsForProject
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Save these options for this particular project.
'---------------------------------------------------------------------------------------
'
Public Sub SaveOptionsForProject()
    Me.SaveOptionsToFile Me.GetExportFolder & cstrOptionsFilename
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadOptionsFromFile
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Load matching options from json file.
'---------------------------------------------------------------------------------------
'
Public Sub LoadOptionsFromFile(strFile As String)

    Dim dOptions As Scripting.Dictionary
    Dim varOption As Variant
    Dim strKey As String
    
    If FSO.FileExists(strFile) Then
        ' Read file contents
        With FSO.OpenTextFile(strFile)
            Set dOptions = modJsonConverter.ParseJson(.ReadAll)("Options")
            .Close
        End With
        If Not dOptions Is Nothing Then
            ' Attempt to set any matching options in this class.
            For Each varOption In m_colOptions
                strKey = CStr(varOption)
                If dOptions.Exists(strKey) Then
                    ' Set class property with value read from file.
                    Select Case strKey
                        Case "TablesToExportData"
                            Set Me.TablesToExportData = dOptions(strKey)
                        Case Else
                            ' Regular top-level properties
                            CallByName Me, strKey, VbLet, dOptions(strKey)
                    End Select
                End If
            Next varOption
        End If
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadProjectOptions
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Loads the project options from a saved file. (If any)
'---------------------------------------------------------------------------------------
'
Public Sub LoadProjectOptions()
    LoadOptionsFromFile Me.GetExportFolder & cstrOptionsFilename
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadDefaultOptions
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Load in the default options from the saved json file in the addin path.
'---------------------------------------------------------------------------------------
'
Public Sub LoadDefaultOptions()
    LoadOptionsFromFile CodeProject.Path & "\" & FSO.GetBaseName(CodeProject.Name) & ".json"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PrintOptionsToDebugWindow
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Output the current options.
'---------------------------------------------------------------------------------------
'
Public Sub PrintOptionsToDebugWindow()
    Debug.Print modJsonConverter.ConvertToJson(SerializeOptions, JSON_WHITESPACE)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetExportFolder
' Author    : Adam Waller
' Date      : 4/14/2020
' Purpose   : Returns the actual export folder, even if a path hasn't been defined.
'---------------------------------------------------------------------------------------
'
Public Function GetExportFolder() As String
    If Me.ExportFolder = vbNullString Then
        ' Build default path using project name
        GetExportFolder = CurrentProject.FullName & ".src\"
    Else
        ' This should be an absolute path, not a relative one.
        GetExportFolder = Me.ExportFolder
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : SerializeOptions
' Author    : Adam Waller
' Date      : 2/12/2020
' Purpose   : Serializes Options into a dictionary array for saving to file as JSON.
'---------------------------------------------------------------------------------------
'
Private Function SerializeOptions() As Scripting.Dictionary

    Dim dOptions As Scripting.Dictionary
    Dim dInfo As Scripting.Dictionary
    Dim dWrapper As Scripting.Dictionary
    Dim varOption As Variant
    Dim strOption As String
    Dim strBit As String
    
    Set dOptions = New Scripting.Dictionary
    Set dInfo = New Scripting.Dictionary
    Set dWrapper = New Scripting.Dictionary
    
    ' Add some header information (For debugging or upgrading)
    #If Win64 Then
        strBit = " 64-bit"
    #Else
        strBit = " 32-bit"
    #End If
    dInfo.Add "AddinVersion", AppVersion
    dInfo.Add "AccessVersion", Application.Version & strBit
    dInfo.Add "Hash", Encrypt(CodeProject.Name)
    
    For Each varOption In m_colOptions
        strOption = CStr(varOption)
        ' Simulate reflection to serialize properties
        dOptions.Add CStr(strOption), CallByName(Me, strOption, VbGet)
    Next varOption
    
    'Set SerializeOptions = New Scripting.Dictionary
    Set dWrapper("Info") = dInfo
    Set dWrapper("Options") = dOptions
    Set SerializeOptions = dWrapper
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTableExportFormatName
' Author    : Adam Waller
' Date      : 4/17/2020
' Purpose   : Return the name used to read and write to the JSON options files.
'---------------------------------------------------------------------------------------
'
Public Function GetTableExportFormatName(intFormat As eTableDataExportFormat) As String
    Select Case intFormat
        Case etdTabDelimited:   GetTableExportFormatName = "Tab Delimited"
        Case etdXML:            GetTableExportFormatName = "XML Format"
        Case Else:              GetTableExportFormatName = vbNullString
    End Select
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTableExportFormat
' Author    : Adam Waller
' Date      : 4/17/2020
' Purpose   : Translate the table export format key to the corresponding enum value.
'---------------------------------------------------------------------------------------
'
Public Function GetTableExportFormat(strKey As String) As eTableDataExportFormat
    Dim intFormat As eTableDataExportFormat
    Dim strName As String
    For intFormat = etdNoData To eTableDataExportFormat.[_Last]
        strName = Me.GetTableExportFormatName(intFormat)
        If strName = strKey Then
            GetTableExportFormat = intFormat
            Exit For
        End If
    Next intFormat
End Function

'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 2/12/2020
' Purpose   : Load default values when creating class
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    
    ' Load list of property names for reflection type behavior.
    With m_colOptions
        .Add "ExportFolder"
        .Add "ShowDebug"
        .Add "UseFastSave"
        .Add "SavePrintVars"
        .Add "SaveQuerySQL"
        .Add "SaveTableSQL"
        .Add "StripPublishOption"
        .Add "AggressiveSanitize"
        .Add "TablesToExportData"
        .Add "RunBeforeExport"
        .Add "RunAfterExport"
        .Add "RunAfterBuild"
        .Add "KeyName"
    End With
    
    ' Load default values
    Me.LoadDefaults
    
    ' Other run-time options
    JsonOptions.AllowUnicodeChars = True
    
End Sub