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
Public IncludeVBE As Boolean
Public UseFastSave As Boolean
Public SavePrintVars As Boolean
Public SaveQuerySQL As Boolean
Public SaveTableSQL As Boolean
Public StripPublishOption As Boolean
Public AggressiveSanitize As Boolean
Public TablesToSave As New Collection

Private m_colOptions As New Collection


'---------------------------------------------------------------------------------------
' Procedure : LoadDefaults
' Author    : Adam Waller
' Date      : 2/12/2020
' Purpose   : Loads the default values. Define system defaults here.
'---------------------------------------------------------------------------------------
'
Public Sub LoadDefaults()

    With Me
        .ExportFolder = ""
        .ShowDebug = False
        .IncludeVBE = False
        .UseFastSave = True
        .SavePrintVars = True
        .SaveQuerySQL = True
        .SaveTableSQL = True
        .StripPublishOption = True
        .AggressiveSanitize = True
        Set .TablesToSave = New Collection
        ' Save specific tables by default
        SaveTableIfExists "USysRibbons", "USysRegInfo"
    End With
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveTableIfExists
' Author    : Adam Waller
' Date      : 2/12/2020
' Purpose   : Function to add table to save list if it exists in the current database
'---------------------------------------------------------------------------------------
'
Private Sub SaveTableIfExists(ParamArray varTableNames() As Variant)
    
    Dim tdf As Access.AccessObject
    Dim varTable As Variant
    
    For Each tdf In CurrentData.AllTables
        For Each varTable In varTableNames
            If tdf.Name = varTable Then
                Me.TablesToSave.Add CStr(varTable)
            End If
        Next varTable
    Next tdf
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveOptionsToFile
' Author    : Adam Waller
' Date      : 2/12/2020
' Purpose   : Save the loaded Options to a file in JSON format
'---------------------------------------------------------------------------------------
'
Public Sub SaveOptionsToFile(strPath As String)
    WriteFile modJsonConverter.ConvertToJson(SerializeOptions, 2) & vbCrLf, strPath
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
    Dim varItem As Variant
    
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
                        Case "TablesToSave"
                            Set Me.TablesToSave = New Collection
                            For Each varItem In dOptions(strKey)
                                Me.TablesToSave.Add CStr(varItem)
                            Next varItem
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
    Debug.Print modJsonConverter.ConvertToJson(SerializeOptions, 2)
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
        .Add "IncludeVBE"
        .Add "UseFastSave"
        .Add "SavePrintVars"
        .Add "SaveQuerySQL"
        .Add "SaveTableSQL"
        .Add "StripPublishOption"
        .Add "AggressiveSanitize"
        .Add "TablesToSave"
    End With
    
    ' Load default values
    Me.LoadDefaults
    
End Sub