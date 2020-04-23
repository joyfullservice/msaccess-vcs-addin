Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : IDbComponent (Abstract Class)
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This class interface defines the standard properties for exporting and
'           : importing database objects for version control. This class should be
'           : implemented into the classes defined for each object type to be exported
'           : and imported as a part of the database source code.
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Public Sub Export()
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Public Sub Import(strFile As String)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Public Function GetAllFromDB(Optional cOptions As clsOptions) As Collection
End Function


'---------------------------------------------------------------------------------------
' Procedure : ModifiedDate
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'           : If the modified date cannot be determined (such as application
'           : properties) then this function will return 0.
'---------------------------------------------------------------------------------------
'
Public Function ModifiedDate() As Date
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'           : Based on the export folder defined in options.
'---------------------------------------------------------------------------------------
'
Public Function GetFileList() As Collection
End Function


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Public Property Get ComponentType() As eDatabaseComponentType
End Property


'---------------------------------------------------------------------------------------
' Procedure : Upgrade
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Run any version specific upgrade processes before importing.
'---------------------------------------------------------------------------------------
'
Public Sub Upgrade()
End Sub