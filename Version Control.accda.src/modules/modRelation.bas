Option Explicit
Option Compare Database
Option Private Module


'---------------------------------------------------------------------------------------
' Procedure : ExportRelation
' Author    : Adam Waller
' Date      : 1/24/2019
' Purpose   : Exports the database table relationships
'---------------------------------------------------------------------------------------
'
Public Sub ExportRelation(rel As Relation, strFile As String)

    Dim cData As New clsConcat
    Dim fld As DAO.Field
    
    With cData
        .Add rel.Attributes 'RelationAttributeEnum
        .Add rel.Name
        .Add rel.table
        .Add rel.foreignTable
        For Each fld In rel.Fields
            .Add "Field = Begin"
            .Add fld.Name
            .Add fld.ForeignName
            .Add "End"
        Next
    End With
    WriteFile cData.GetStr, strFile
    
End Sub


Public Sub ImportRelation(filePath As String)

    Dim InFile As Scripting.TextStream
    Set InFile = FSO.OpenTextFile(filePath, 1)
    
    Dim rel As New Relation
    rel.Attributes = InFile.ReadLine
    rel.Name = InFile.ReadLine
    rel.table = InFile.ReadLine
    rel.foreignTable = InFile.ReadLine
    Dim f As Object ' Field
    Do Until InFile.AtEndOfStream
        If "Field = Begin" = InFile.ReadLine Then
            'Set f = New Field
            Set f = CreateObject("ADODB.Field")
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
        GetRelationFileName = GetSafeFileName(CStr(Split(strName, "].")(1)))
    Else
        GetRelationFileName = GetSafeFileName(strName)
    End If

End Function