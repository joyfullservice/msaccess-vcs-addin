Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


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
' Purpose   : Save the loaded options to a file in JSON format
'---------------------------------------------------------------------------------------
'
Public Sub SaveOptionsToFile(strPath As String)

End Sub


Public Sub PrintOptions()
    Debug.Print modJsonConverter.ConvertToJson(SerializeOptions, 2)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SerializeOptions
' Author    : Adam Waller
' Date      : 2/12/2020
' Purpose   : Serializes options into a dictionary array for saving to file as JSON.
'---------------------------------------------------------------------------------------
'
Private Function SerializeOptions() As Scripting.Dictionary

    Dim dOptions As New Scripting.Dictionary
    Dim dWrapper As New Scripting.Dictionary
    Dim varOption As Variant
    Dim strOption As String
    
    For Each varOption In m_colOptions
        strOption = CStr(varOption)
        ' Simulate reflection to serialize properties
        dOptions.Add CStr(strOption), CallByName(Me, strOption, VbGet)
    Next varOption
    
    'Set SerializeOptions = New Scripting.Dictionary
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