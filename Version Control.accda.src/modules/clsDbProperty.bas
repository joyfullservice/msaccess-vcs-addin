Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This class extends the IDbComponent class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private m_Property As DAO.Property
Private m_Options As clsOptions
'Private m_Count As Long (uncomment if needed)

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export()
    
    Dim strFile As String
    Dim strTempFile As String

    ' Check for existing file
    strFile = IDbComponent_SourceFile
    If FSO.FileExists(strFile) Then Kill strFile
    ExportPropertiesJson
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional cOptions As clsOptions) As Collection
    
    Dim prp As DAO.Property
    Dim cProp As IDbComponent

    ' Use parameter options if provided.
    If Not cOptions Is Nothing Then Set IDbComponent_Options = cOptions

    Set IDbComponent_GetAllFromDB = New Collection
    For Each prp In CurrentDb.Properties
        Set cProp = New clsDbProperty
        Set cProp.DbObject = prp
        Set cProp.Options = IDbComponent_Options
        IDbComponent_GetAllFromDB.Add cProp, prp.Name
    Next prp
        
End Function


'---------------------------------------------------------------------------------------
' Procedure : ExportProperties
' Author    : Adam Waller
' Date      : 1/24/2019
' Purpose   : Export database properties to a CSV
'---------------------------------------------------------------------------------------
'
Public Sub ExportPropertiesJson()
    
    Dim prp As DAO.Property
    Dim dCollection As Scripting.Dictionary
    Dim dItem As Scripting.Dictionary
    
    Set dCollection = New Scripting.Dictionary
    
    ' Loop through all properties
    For Each prp In CurrentDb.Properties
        Select Case prp.Name
            Case "Name"         ' Ignore file name property, since this could contain PI and can't be set anyway.
            Case "Connection"   ' Connection object for ODBCDirect workspaces. Not needed.
            Case Else
                Set dItem = New Scripting.Dictionary
                dItem.Add "Value", prp.Value
                dItem.Add "Type", prp.Type
                dCollection.Add prp.Name, dItem
        End Select
    Next prp
    
    ' Write to file
    WriteJsonFile Me, dCollection, IDbComponent_SourceFile, "Database Properties (DAO)"
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ExportProperties
' Author    : Adam Waller
' Date      : 1/24/2019
' Purpose   : Export database properties to a CSV
'---------------------------------------------------------------------------------------
'
Public Sub ExportProperties(strFolder As String, cOptions As clsOptions)
    
    Const UnitSeparator = "?"  ' Chr(31) INFORMATION SEPARATOR ONE
    
    Dim cData As New clsConcat
    Dim intCnt As Integer
    Dim objParent As Object
    Dim prp As Object
    
    Set objParent = ThisProjectDB
    
    On Error Resume Next
    For Each prp In objParent.Properties
        Select Case prp.Name
            Case "Name"
                ' Ignore file name property, since this could contain PI and can't be set anyway.
            Case Else
                With cData
                    .Add prp.Name
                    .Add UnitSeparator
                    .Add prp.Value
                    .Add UnitSeparator
                    .Add prp.Type
                    .Add vbCrLf
                End With
                
                intCnt = intCnt + 1
        End Select
    Next prp
    
    If Err Then Err.Clear
    On Error GoTo 0
    
    ' Write to file
    WriteFile cData.GetStr, strFolder & "properties.txt"
    
    ' Display summary.
    If cOptions.ShowDebug Then
        Log "[" & intCnt & "] database properties exported."
    Else
        Log "[" & intCnt & "]"
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Module    : ImportProperties
' Author    : Adam Kauffman
' Date      : 2020-01-10
' Purpose   : Import database properties from the exported source
'---------------------------------------------------------------------------------------

' Import database properties from a text file, true=SUCCESS
Private Function ImportProperties(ByVal sourcePath As String, Optional ByRef appInstance As Application) As Boolean
    If appInstance Is Nothing Then Set appInstance = Application.Application
      
    Const UnitSeparator = "?"  ' Chr(31) INFORMATION SEPARATOR ONE

    Dim propertiesFile As String
    propertiesFile = Dir(sourcePath & "properties.txt")
    If Len(propertiesFile) = 0 Then ' File not foud
        ImportProperties = False
        Exit Function
    End If
    
    Debug.Print PadRight("Importing Properties...", cintPad);
    
    Dim thisDb As Object
    Set thisDb = ThisProjectDB(appInstance)
   
    Dim inputFile As Object
    Set inputFile = FSO.OpenTextFile(sourcePath & propertiesFile, ForReading)
    
    Dim propertyCount As Long
    On Error GoTo ErrorHandler
    Do Until inputFile.AtEndOfStream
        Dim recordUnit() As String
        recordUnit = Split(inputFile.ReadLine, UnitSeparator)
        If UBound(recordUnit) > 1 Then ' Looks like a valid entry
            propertyCount = propertyCount + 1
            
            Dim propertyName As String
            Dim propertyValue As Variant
            Dim propertyType As Long
            propertyName = recordUnit(0)
            propertyValue = recordUnit(1)
            propertyType = recordUnit(2)
            
            SetProperty propertyName, propertyValue, thisDb, propertyType
        End If
    Loop
    
ErrorHandler:
    If Err.Number > 0 Then
        If Err.Number = 3001 Then
            ' Invalid argument; means that this property cannot be set by code.
        ElseIf Err.Number = 3032 Then
            ' Cannot perform this operation; means that this property cannot be set by code.
        ElseIf Err.Number = 3259 Then
            ' Invalid field data type; means that the property was not found, use create.
        ElseIf Err.Number = 3251 Then
            ' Operation is not supported for this type of object; means that this property cannot be set by code.
        Else
            Debug.Print " Error: " & Err.Number & " " & Err.Description
        End If
        
        Err.Clear
        Resume Next
    End If
    
    On Error GoTo 0
    
    Debug.Print "[" & propertyCount & "]"
    inputFile.Close
    Set inputFile = Nothing
    ImportProperties = True

End Function

' SetProperty() requires either propertyType is set explicitly OR
'   propertyValue has a valid value and type for a new property to be created.
Private Sub SetProperty(ByVal propertyName As String, ByVal propertyValue As Variant, _
                       Optional ByRef thisDb As Object, _
                       Optional ByVal propertyType As Integer = -1)
                       
    If thisDb Is Nothing Then Set thisDb = ThisProjectDB
    
    Dim newProperty As Property
    Set newProperty = GetProperty(propertyName, thisDb)
    If Not newProperty Is Nothing Then
        If newProperty.Value <> propertyValue Then newProperty.Value = propertyValue
    Else ' Property not found
        If propertyType = -1 Then propertyType = DBVal(varType(propertyValue)) ' Guess the type (Good luck)
        Set newProperty = thisDb.CreateProperty(propertyName, propertyType, propertyValue)
        thisDb.Properties.Append newProperty
    End If
End Sub

' Returns nothing upon Error
Private Function GetProperty(ByVal propertyName As String, _
                            Optional ByRef thisDb As Object) As Property
                            
    Const PropertyNotFound As Integer = 3270
    If thisDb Is Nothing Then Set thisDb = ThisProjectDB
    
    On Error GoTo Err_PropertyExists
    Set GetProperty = thisDb.Properties(propertyName)

    Exit Function
     
Err_PropertyExists:
    If Err.Number <> PropertyNotFound Then
        Debug.Print "Error getting property: " & propertyName & vbNewLine & Err.Number & " " & Err.Description
    End If
    
    Err.Clear
End Function

'   HERE BE DRAGONS
' Return db property type that closely matches VBA varible type
Private Function DBVal(ByVal intVBVal As Integer) As Integer
    Const TypeVBToDB As String = "\2|3\3|4\4|6\5|7\6|5" & _
                                 "\7|8\8|10\11|1\14|20\17|2"
    Dim intX As Integer
    intX = InStr(1, TypeVBToDB, "\" & intVBVal & "|")
    DBVal = Val(Mid$(TypeVBToDB, intX + Len(intVBVal) + 2))
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Collection
    IDbComponent_GetFileList.Add IDbComponent_SourceFile
End Function


'---------------------------------------------------------------------------------------
' Procedure : ClearOrphanedSourceFiles
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Remove any source files for objects not in the current database.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_ClearOrphanedSourceFiles() As Variant
End Function


'---------------------------------------------------------------------------------------
' Procedure : DateModified
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'           : If the modified date cannot be determined (such as application
'           : properties) then this function will return 0.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_DateModified() As Date
    ' Modified date unknown.
    IDbComponent_DateModified = 0
End Function


'---------------------------------------------------------------------------------------
' Procedure : SourceModified
' Author    : Adam Waller
' Date      : 4/27/2020
' Purpose   : The date/time the source object was modified. In most cases, this would
'           : be the date/time of the source file, but it some cases like SQL objects
'           : the date can be determined through other means, so this function
'           : allows either approach to be taken.
'---------------------------------------------------------------------------------------
'
Public Function IDbComponent_SourceModified() As Date
    If FSO.FileExists(IDbComponent_SourceFile) Then IDbComponent_SourceModified = FileDateTime(IDbComponent_SourceFile)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "db properties"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = IDbComponent_Options.GetExportFolder
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    IDbComponent_Name = "Database Properties (DAO)"
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    IDbComponent_SourceFile = IDbComponent_BaseFolder & "dbs-properties.json"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count() As Long
    ' Subtract any named exceptions (such as the Name property)
    IDbComponent_Count = CurrentDb.Properties.Count - 1
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbDbsProperty
End Property


'---------------------------------------------------------------------------------------
' Procedure : Upgrade
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Run any version specific upgrade processes before importing.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Upgrade()
    ' No upgrade needed.
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Options
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : Return or set the options being used in this context.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Options() As clsOptions
    If m_Options Is Nothing Then Set m_Options = LoadOptions
    Set IDbComponent_Options = m_Options
End Property
Private Property Set IDbComponent_Options(ByVal RHS As clsOptions)
    Set m_Options = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_Property
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_Property = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Returns true if the export of all items is done as a single file instead
'           : of individual files for each component. (I.e. properties, references)
'---------------------------------------------------------------------------------------
'
Public Property Get IDbComponent_SingleFile() As Boolean
    IDbComponent_SingleFile = True
End Property


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Helps us know whether we have already counted the tables.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    'm_Count = -1
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Parent
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Return a reference to this class as an IDbComponent. This allows you
'           : to reference the public methods of the parent class without needing
'           : to create a new class object.
'---------------------------------------------------------------------------------------
'
Public Property Get Parent() As IDbComponent
    Set Parent = Me
End Property