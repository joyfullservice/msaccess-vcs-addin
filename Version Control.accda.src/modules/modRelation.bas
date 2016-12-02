Option Compare Database
Option Private Module
Option Explicit


Public Sub ExportRelation(rel As Relation, filePath As String)

    Dim FSO, OutFile As Object ' Scripting.FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set OutFile = FSO.CreateTextFile(filePath, True)

    OutFile.WriteLine rel.Attributes 'RelationAttributeEnum
    OutFile.WriteLine rel.Name
    OutFile.WriteLine rel.table
    OutFile.WriteLine rel.foreignTable
    Dim f As Field
    For Each f In rel.Fields
        OutFile.WriteLine "Field = Begin"
        OutFile.WriteLine f.Name
        OutFile.WriteLine f.ForeignName
        OutFile.WriteLine "End"
    Next
    OutFile.Close

End Sub


Public Sub ImportRelation(filePath As String)
    Dim FSO, InFile As Object ' Scripting.FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set InFile = FSO.OpenTextFile(filePath, 1)
    
    Dim rel As New Relation
    rel.Attributes = InFile.ReadLine
    rel.Name = InFile.ReadLine
    rel.table = InFile.ReadLine
    rel.foreignTable = InFile.ReadLine
    Dim f As Field
    Do Until InFile.AtEndOfStream
        If "Field = Begin" = InFile.ReadLine Then
            Set f = New Field
            f.Name = InFile.ReadLine
            f.ForeignName = InFile.ReadLine
            If "End" <> InFile.ReadLine Then
                Set f = Nothing
                Err.Raise 40000, "ImportRelation", "Missing 'End' for a 'Begin' in " & filePath
            End If
            rel.Fields.Append f
        End If
    Loop
    
    InFile.Close
    
    CurrentDb.Relations.Append rel

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetRelationFileName
' Author    : Adam Waller
' Date      : 6/4/2015
' Purpose   : Build file name based on relation name, including support for linked
'           : tables that would put a slash in the relation name.
'           : (Strips the link path from the table name)
'---------------------------------------------------------------------------------------
'
Public Function GetRelationFileName(objRelation As Relation) As String

    Dim strName As String
    
    strName = objRelation.Name
    
    If InStr(1, strName, "].") > 0 Then
        ' Need to remove path to linked file
        GetRelationFileName = Split(strName, "].")(1)
    Else
        GetRelationFileName = strName
    End If

End Function