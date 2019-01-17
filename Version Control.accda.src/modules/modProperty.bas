Option Explicit
Option Compare Database
Option Private Module


' Import database properties from a text file, true=SUCCESS
Public Function ImportProperties(obj_path As String) As Boolean

    Dim fso As New Scripting.FileSystemObject
    Dim InFile As Scripting.TextStream
    Dim strLine As String
    Dim Item() As String
    Dim GUID As String
    Dim Major As Long
    Dim Minor As Long
    Dim fileName As String
    Dim refName As String
    Dim objParent As Object
    Dim strVal As String
    Dim prp As Object
    
    fileName = Dir(obj_path & "properties.txt")
    If Len(fileName) = 0 Then
        ImportProperties = False
        Exit Function
    End If
    
    Set InFile = fso.OpenTextFile(obj_path & fileName, ForReading)

    Set objParent = CodeDb
    If CodeProject.ProjectType = acADP Then Set objParent = CurrentProject

    'On Error Resume Next
    Do Until InFile.AtEndOfStream
        strLine = InFile.ReadLine
        Item = Split(strLine, "=")
        If UBound(Item) > 0 Then ' Looks like a valid entry
            ' Set property in database
            Set prp = objParent.Properties(Item(0))
            strVal = Mid(strLine, Len(Item(0)) + 2)
            If prp.Value <> strVal Then
                ' Different property. Attempt to set.
                objParent.Properties(Item(0)) = strVal
            End If
        End If
    Loop
    If Err Then Err.Clear
    On Error GoTo 0
    
    InFile.Close
    Set InFile = Nothing
    Set fso = Nothing
    ImportProperties = True

End Function


' Export database properties to a CSV
Public Sub ExportProperties(obj_path As String)
    
    Dim fso As New Scripting.FileSystemObject
    Dim OutFile As Scripting.TextStream
    Dim obj_count As Integer
    Dim objParent As Object
    Dim prp As Object
    
    Set OutFile = fso.CreateTextFile(obj_path & "properties.txt", True)
    
    ' Save list of properties set in current database.
    Set objParent = CodeDb
    If CodeProject.ProjectType = acADP Then Set objParent = CurrentProject
    
    On Error Resume Next
    For Each prp In objParent.Properties
        ' Ignore file name property, since this could contain PI and can't be set anyway.
        If prp.Name <> "Name" Then
            OutFile.WriteLine prp.Name & "=" & prp.Value
            obj_count = obj_count + 1
        End If
    Next prp
    If Err Then Err.Clear
    On Error GoTo 0
    
    OutFile.Close

    If ShowDebugInfo Then
        Debug.Print "[" & obj_count & "] database properties exported."
    Else
        Debug.Print "[" & obj_count & "]"
    End If
    
    Set OutFile = Nothing
    Set fso = Nothing
    
End Sub