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


' Objects and collections need to be set with wrappers.
Private m_TablesToSaveData As New Collection


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
Public Property Get SelectionSourceFile(Optional UseVBEFile As Boolean = True) As String
    
    Dim strPath As String
    Dim strName As String
    
    If ProjectIsSelected Then
        ' Path to root project folder
        strPath = Me.ExportBaseFolder
        ' Trim trailing slash
        If Right(strPath, 1) = "\" Then strPath = Left(strPath, Len(strPath) - 1)
    Else
        ' Get correct file extension and path
        strName = VBE.SelectedVBComponent.Name
        If UseVBEFile Then
            strName = strName & GetVBEExtByType(VBE.SelectedVBComponent)
            strPath = Me.ExportBaseFolder & "VBE\" & strName
            ' Fall back to database object if VBE file not found
            If Dir(strPath) = "" Then strPath = ""
        End If
        If strPath = "" Then
            ' Remove prefixes
            strName = Replace(strName, "Form_", "", , 1)
            strName = Replace(strName, "Report_", "", , 1)
            strName = strName & ".bas"
            strPath = Me.ExportBaseFolder & strName
            If Dir(strPath) = "" Then strPath = ""
        End If
    End If
    
    ' Return path to source file
    If strPath = "" Then
        MsgBox "Path to exported source file not found. Please export first.", vbExclamation
    Else
        ' Add quotes to path
        SelectionSourceFile = """" & strPath & """"
    End If

End Property


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
Public Property Get HasRequiredSoftware(blnWarnUser As Boolean) As Boolean
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