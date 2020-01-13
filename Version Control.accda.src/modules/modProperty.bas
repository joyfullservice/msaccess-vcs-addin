Option Explicit
Option Compare Database
Option Private Module

Private Const UnitSeparator = "?"  ' Chr(31) INFORMATION SEPARATOR ONE

Public Function ThisProjectDB(Optional ByRef appInstance As Application) As Object
    If appInstance Is Nothing Then Set appInstance = Application.Application
    If CurrentProject.ProjectType = acMDB Then
        Set ThisProjectDB = appInstance.CurrentDb
    Else  ' ADP project
        Set ThisProjectDB = appInstance.CurrentProject
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : ExportProperties
' Author    : Adam Waller
' Date      : 1/24/2019
' Purpose   : Export database properties to a CSV
'---------------------------------------------------------------------------------------
'
Public Sub ExportProperties(strFolder As String, cModel As IVersionControl)
    
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
    If cModel.ShowDebug Then
        cModel.Log "[" & intCnt & "] database properties exported."
    Else
        cModel.Log "[" & intCnt & "]"
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Module    : ImportProperties
' Author    : Adam Kauffman
' Date      : 2020-01-10
' Purpose   : Import database properties from the exported source
'---------------------------------------------------------------------------------------

' Import database properties from a text file, true=SUCCESS
Public Function ImportProperties(ByVal sourcePath As String, Optional ByRef appInstance As Application) As Boolean
    If appInstance Is Nothing Then Set appInstance = Application.Application
      
    Dim propertiesFile As String
    propertiesFile = Dir(sourcePath & "properties.txt")
    If Len(propertiesFile) = 0 Then ' File not foud
        ImportProperties = False
        Exit Function
    End If
    
    Debug.Print PadRight("Importing Properties...", 24);
    
    Dim thisDB As Object
    Set thisDB = ThisProjectDB(appInstance)
   
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
            
            SetProperty propertyName, propertyValue, thisDB, propertyType
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
Public Sub SetProperty(ByVal propertyName As String, ByVal propertyValue As Variant, _
                       Optional ByRef thisDB As Object, _
                       Optional ByVal propertyType As Integer = -1)
                       
    If thisDB Is Nothing Then Set thisDB = ThisProjectDB
    
    Dim newProperty As Property
    Set newProperty = GetProperty(propertyName, thisDB)
    If Not newProperty Is Nothing Then
        If newProperty.Value <> propertyValue Then newProperty.Value = propertyValue
    Else ' Property not found
        If propertyType = -1 Then propertyType = DBVal(varType(propertyValue)) ' Guess the type (Good luck)
        Set newProperty = thisDB.CreateProperty(propertyName, propertyType, propertyValue)
        thisDB.Properties.Append newProperty
    End If
End Sub

' Returns nothing upon Error
Public Function GetProperty(ByVal propertyName As String, _
                            Optional ByRef thisDB As Object) As Property
                            
    Const PropertyNotFound As Integer = 3270
    If thisDB Is Nothing Then Set thisDB = ThisProjectDB
    
    On Error GoTo Err_PropertyExists
    Set GetProperty = thisDB.Properties(propertyName)

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