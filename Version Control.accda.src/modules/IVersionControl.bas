Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Database


' Simple properties can be set here
Public ExportBaseFolder As String
Public ShowDebug As Boolean
Public IncludeVBE As Boolean
Public FastSave As Boolean
Public SavePrintVars As Boolean
Public SaveQuerySQL As Boolean
Public SaveTableSQL As Boolean
Public AggressiveSanitize As Boolean
Public StripPublishOption As Boolean


' Objects and collections need to be set with wrappers.
Private m_TablesToSaveData As New Collection
Private m_Log As New clsConcat


'---------------------------------------------------------------------------------------
' Procedure : TablesToSaveData
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Tables where we would like to save data during export
'---------------------------------------------------------------------------------------
'
Public Property Get TablesToSaveData() As Collection
    Set TablesToSaveData = m_TablesToSaveData
End Property


'---------------------------------------------------------------------------------------
' Procedure : SelectionSourceFile
' Author    : Adam Waller
' Date      : 6/3/2015
' Purpose   : Returns the full path to the selected component
'---------------------------------------------------------------------------------------
'
Public Property Get SelectionSourceFile(Optional ByRef UseVBEFile As Boolean = True) As String
    
    Dim strPath As String
    Dim strName As String
    
    If ProjectIsSelected Then
        ' Path to root project folder
        strPath = StripSlash(Me.ExportBaseFolder)
    Else
        ' Get correct file extension and path
        strName = VBE.SelectedVBComponent.Name
        If UseVBEFile Then
            strName = strName & GetVBEExtByType(VBE.SelectedVBComponent)
            strPath = Me.ExportBaseFolder & "VBE\" & strName
            ' Fall back to database object if VBE file not found
            If Dir(strPath) = vbNullString Then strPath = vbNullString
        End If
        If strPath = vbNullString Then
            ' Remove prefixes
            strName = Replace(strName, "Form_", vbNullString, , 1)
            strName = Replace(strName, "Report_", vbNullString, , 1)
            strName = strName & ".bas"
            strPath = Me.ExportBaseFolder & strName
            If Dir(strPath) = vbNullString Then strPath = vbNullString
        End If
    End If
    
    ' Return path to source file
    If strPath = vbNullString Then
        MsgBox "Path to exported source file not found. Please export first.", vbExclamation
    Else
        ' Add quotes to path
        SelectionSourceFile = """" & strPath & """"
    End If

End Property


'---------------------------------------------------------------------------------------
' Procedure : Log
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Add a log file entry.
'---------------------------------------------------------------------------------------
'
Public Sub Log(ByRef strText As String, Optional ByRef blnPrint As Boolean = True, Optional ByRef blnNextOutputOnNewLine As Boolean = True)

    Static dblLastLog As Double
    
    m_Log.Add strText
    If blnPrint Then
        If blnNextOutputOnNewLine Then
            ' Create new line
            Debug.Print strText
        Else
            ' Continue next printout on this line.
            Debug.Print strText;
        End If
    End If
    
    If blnNextOutputOnNewLine Then m_Log.Add vbCrLf
    
    ' Allow an update to the screen every second.
    ' (This keeps the aplication from an apparent hang while
    '  running intensive export processes.)
    If dblLastLog + 1 < Timer Then
        DoEvents
        dblLastLog = Timer
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveLogFile
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Saves the log data to a file, and resets the log buffer.
'---------------------------------------------------------------------------------------
'
Public Sub SaveLogFile(ByRef strPath As String)
    WriteFile m_Log.GetStr, strPath
    Set m_Log = New clsConcat
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 5/18/2015
' Purpose   : Save the selected item or project, exporting files to the sources folder
'---------------------------------------------------------------------------------------
'
Public Sub Export()
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportAll
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Export the entire project
'---------------------------------------------------------------------------------------
'
Public Sub ExportAll()
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Commit
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Commit the changes to the repository
'---------------------------------------------------------------------------------------
'
Public Sub Commit()
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Diff
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Diff the changes on the selected item (or project)
'---------------------------------------------------------------------------------------
'
Public Sub Diff()
End Sub


'---------------------------------------------------------------------------------------
' Procedure : HasRequiredSoftware
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Make sure the user has the software installed to use the VCS.
'---------------------------------------------------------------------------------------
'
Public Property Get HasRequiredSoftware(ByRef blnWarnUser As Boolean) As Boolean
End Property


'---------------------------------------------------------------------------------------
' Procedure : Terminate
' Author    : Adam Waller
' Date      : 6/2/2015
' Purpose   : Manually call the terminate event.
'           : This is helpful when you are dealing with circular references
'           : in class objects. http://dailydoseofexcel.com/archives/2007/12/28/terminating-dependent-classes/
'---------------------------------------------------------------------------------------
'
Public Sub Terminate()
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 1/24/2019
' Purpose   : Set up default values. (Can be overriden later)
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    ExportBaseFolder = CurrentProject.Path & "\" & CurrentProject.Name & ".src\"
    AggressiveSanitize = True
    StripPublishOption = True
    IncludeVBE = False
    FastSave = True
    SavePrintVars = False
    ShowDebug = False
End Sub