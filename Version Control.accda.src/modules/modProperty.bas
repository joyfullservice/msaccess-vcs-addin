Option Explicit
Option Compare Database
Option Private Module


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
    
    ' Save list of properties set in current database.
    If CodeProject.ProjectType = acMDB Then
        Set objParent = CurrentDb
    Else
        ' ADP project
        Set objParent = CurrentProject
    End If
    
    On Error Resume Next
    For Each prp In objParent.Properties
        Select Case prp.Name
            Case "Name"
                ' Ignore file name property, since this could contain PI and can't be set anyway.
            Case Else
                With cData
                    .Add prp.Name
                    .Add "="
                    .Add prp.Value
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


' Import database properties from a text file, true=SUCCESS
Public Function ImportProperties(obj_path As String) As Boolean

    Dim InFile As Scripting.TextStream
    Dim strLine As String
    Dim Item() As String
    Dim FileName As String
    Dim objParent As Object
    Dim strVal As String
    Dim prp As Object
    
    FileName = Dir(obj_path & "properties.txt")
    If Len(FileName) = 0 Then
        ImportProperties = False
        Exit Function
    End If
    
    Set InFile = FSO.OpenTextFile(obj_path & FileName, ForReading)

    Set objParent = CurrentDb
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
    ImportProperties = True

End Function