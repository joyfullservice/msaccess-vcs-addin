Option Compare Database
Option Private Module
Option Explicit


' Import database properties from a text file, true=SUCCESS
Public Function ImportProperties(obj_path As String) As Boolean

    Dim FSO, InFile
    Dim strLine As String
    Dim item() As String
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
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set InFile = FSO.OpenTextFile(obj_path & fileName, ForReading)

    Set objParent = CodeDb
    If CodeProject.ProjectType = acADP Then Set objParent = CurrentProject

    'On Error Resume Next
    Do Until InFile.AtEndOfStream
        strLine = InFile.ReadLine
        item = Split(strLine, "=")
        If UBound(item) > 0 Then ' Looks like a valid entry
            ' Set property in database
            Set prp = objParent.Properties(item(0))
            strVal = Mid(strLine, Len(item(0)) + 2)
            If prp.Value <> strVal Then
                ' Different property. Attempt to set.
                objParent.Properties(item(0)) = strVal
            End If
        End If
    Loop
    If Err Then Err.Clear
    On Error GoTo 0
    
    InFile.Close
    Set InFile = Nothing
    Set FSO = Nothing
    ImportProperties = True

End Function


' Export database properties to a CSV
Public Sub ExportProperties(obj_path As String)
    
    Dim FSO As Object ' Scripting.FileSystemObject
    Dim OutFile As Object
    Dim strLine As String
    Dim ref As Reference
    Dim obj_count As Integer
    Dim objParent As Object
    Dim prp As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OutFile = FSO.CreateTextFile(obj_path & "properties.txt", True)
    
    ' Save list of properties set in current database.
    Set objParent = CodeDb
    If CodeProject.ProjectType = acADP Then Set objParent = CurrentProject
    
    On Error Resume Next
    For Each prp In objParent.Properties
        OutFile.WriteLine prp.Name & "=" & prp.Value
        obj_count = obj_count + 1
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
    Set FSO = Nothing
    
End Sub