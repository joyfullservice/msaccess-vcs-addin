Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' Simple properties can be set here
Public ExportBaseFolder As String
Public ShowDebug As Boolean
Public IncludeVBE As Boolean


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