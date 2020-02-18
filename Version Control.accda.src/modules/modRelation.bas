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
        .Add vbCrLf
        .Add rel.Name
        .Add vbCrLf
        .Add rel.Table
        .Add vbCrLf
        .Add rel.ForeignTable
        .Add vbCrLf
        For Each fld In rel.Fields
            .Add "Field = Begin"
            .Add vbCrLf
            .Add fld.Name
            .Add vbCrLf
            .Add fld.ForeignName
            .Add vbCrLf
            .Add "End"
            .Add vbCrLf
        Next
    End With
    WriteFile cData.GetStr, strFile
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ImportRelation
' Author    : Adam Kauffman
' Date      : 02/18/2020
' Purpose   : Import a table relationship
'---------------------------------------------------------------------------------------
'
Public Sub ImportRelation(ByRef filePath As String, Optional ByRef appInstance As Application)
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim thisDb As Database
    Set thisDb = appInstance.CurrentDb
    
    Dim fileLines() As String
    With FSO.OpenTextFile(filePath, IOMode:=ForReading, create:=False, Format:=TristateFalse)
        fileLines = Split(.ReadAll, vbCrLf)
        .Close
    End With
    
    Dim newRelation As Relation
    Set newRelation = thisDb.CreateRelation(fileLines(1), fileLines(2), fileLines(3), fileLines(0))
    
    Dim newField As Field
    Dim thisLine As Long
    For thisLine = 4 To UBound(fileLines)
        If "Field = Begin" = fileLines(thisLine) Then
            thisLine = thisLine + 1
            Set newField = newRelation.CreateField(fileLines(thisLine))  ' Name set here
            thisLine = thisLine + 1
            newField.ForeignName = fileLines(thisLine)
            thisLine = thisLine + 1
            If "End" <> fileLines(thisLine) Then
                Set newField = Nothing
                Err.Raise 40000, "ImportRelation", "Missing 'End' for a 'Begin' in " & filePath
            End If
            
            newRelation.Fields.Append newField
        End If
    Next thisLine
        
    ' Remove conflicting Index entries because adding the relation creates new indexes causing "Error 3284 Index already exists"
    On Error Resume Next
    With thisDb
        .Relations.Delete newRelation.Name  ' Avoid 3012 Relationship already exists
        .TableDefs(newRelation.Table).Indexes.Delete newRelation.Name
        .TableDefs(newRelation.ForeignTable).Indexes.Delete newRelation.Name
    End With
    On Error GoTo ErrorHandler
    
    With thisDb.Relations
        .Append newRelation
    End With
    
ErrorHandler:
    Select Case Err.Number
    Case 0
    Case 3012
        Debug.Print "Relationship already exists: """ & newRelation.Name & """ "
    Case 3284
        Debug.Print "Index already exists for: """ & newRelation.Name & """ "
    Case Else
        Debug.Print "Failed to add: """ & newRelation.Name & """ " & Err.Number & " " & Err.Description
    End Select
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