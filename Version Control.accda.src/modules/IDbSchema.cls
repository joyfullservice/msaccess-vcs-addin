VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDbSchema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : IDbSchema (Abstract Class)
' Author    : Adam Waller
' Date      : 7/14/2023
' Purpose   : This class interface defines the standard functions for exporting and
'           : importing database server objects for version control. This class should
'           : be implemented into the classes defined for each server type used when
'           : exporting object definitions for server-side objects.
'           : NOTE: This is not intended to write to an external database server, but
'           : only to download object definitions that may be related to development
'           : in Microsoft Access databases.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : Initialize
' Author    : Adam Waller
' Date      : 7/14/2023
' Purpose   : Initialize the instance of the server connection
'---------------------------------------------------------------------------------------
'
Public Sub Initialize(dInstance As Dictionary)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ObjectCount
' Author    : Adam Waller
' Date      : 7/21/2023
' Purpose   : Return a count of total objects, or how many changes were found between
'           : the database server objects and the current index file.
'---------------------------------------------------------------------------------------
'
Public Function ObjectCount(blnModifiedOnly As Boolean) As Long
End Function


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 7/14/2023
' Purpose   : Export the database server objects (schema) to source files
'---------------------------------------------------------------------------------------
'
Public Sub Export(blnFullExport As Boolean, Optional strAlternatePath As String)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 7/14/2023
' Purpose   : Base folder for export, (available after initializing)
'---------------------------------------------------------------------------------------
'
Public Property Get BaseFolder() As String
End Property


'---------------------------------------------------------------------------------------
' Procedure : ServerType
' Author    : Adam Waller
' Date      : 7/14/2023
' Purpose   : The type of database server represented by this class
'---------------------------------------------------------------------------------------
'
Public Property Get ServerType() As eDatabaseServerType
End Property


'---------------------------------------------------------------------------------------
' Procedure : TypeDescription
' Author    : Adam Waller
' Date      : 7/14/2023
' Purpose   : Return description for this server type (i.e. Microsoft SQL Server)
'---------------------------------------------------------------------------------------
'
Public Property Get TypeDescription()
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 7/14/2023
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Public Property Get Name() As String
End Property
