Attribute VB_Name = "VCS_Relation"
Option Compare Database

Option Private Module
Option Explicit


Public Sub VCS_ExportRelation(ByVal rel As DAO.Relation, ByVal filePath As String)
    Dim FSO As Object
    Dim OutFile As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OutFile = FSO.CreateTextFile(filePath, overwrite:=True, Unicode:=False)

    OutFile.WriteLine rel.Attributes 'RelationAttributeEnum
    OutFile.WriteLine rel.name
    OutFile.WriteLine rel.table
    OutFile.WriteLine rel.foreignTable
    
    Dim f As DAO.Field
    For Each f In rel.Fields
        OutFile.WriteLine "Field = Begin"
        OutFile.WriteLine f.name
        OutFile.WriteLine f.ForeignName
        OutFile.WriteLine "End"
    Next
    
    OutFile.Close

End Sub

Public Sub VCS_ImportRelation(ByVal filePath As String)
    Dim FSO As Object
    Dim InFile As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set InFile = FSO.OpenTextFile(filePath, iomode:=ForReading, create:=False, Format:=TristateFalse)
    Dim rel As DAO.Relation
    Set rel = New DAO.Relation
    
    rel.Attributes = InFile.ReadLine
    rel.name = InFile.ReadLine
    rel.table = InFile.ReadLine
    rel.foreignTable = InFile.ReadLine
    
    Dim f As DAO.Field
    Do Until InFile.AtEndOfStream
        If "Field = Begin" = InFile.ReadLine Then
            Set f = New DAO.Field
            f.name = InFile.ReadLine
            f.ForeignName = InFile.ReadLine
            If "End" <> InFile.ReadLine Then
                Set f = Nothing
                Err.Raise 40000, "VCS_ImportRelation", "Missing 'End' for a 'Begin' in " & filePath
            End If
            rel.Fields.Append f
        End If
    Loop
    
    InFile.Close
    
    ' Skip if relationship already exists and make a note of it. It was embedded in the table schema.
    On Error GoTo ErrorHandler
    CurrentDb.Relations.Append rel
    
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 3012    ' Relationship already exists
            Debug.Print "Skipped: """ & rel.Name & """ ";
            Resume Next    ' Skip it and move on
        Case Else
            Resume Next    ' Move on anyways
    End Select
End Sub
