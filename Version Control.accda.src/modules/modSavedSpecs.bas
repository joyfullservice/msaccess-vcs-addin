Option Compare Database
Option Private Module
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : ExportSpecs
' Author    : Adam Waller
' Date      : 2/6/2019
' Purpose   : Export the import/export specifications into a hybrid file format that
'           : includes the name, description, and XML content.
'---------------------------------------------------------------------------------------
'
Public Sub ExportSpecs(strSourcePath As String, cOptions As clsOptions)
    
    Dim strFolder As String
    Dim oSpec As ImportExportSpecification
    Dim cData As New clsConcat
    Dim strXML As String
    Dim strName As String
    Dim strFile As String
    
    ' Clear any existing files
    ClearFilesByExtension strSourcePath & "\importspecs", "spec"
    ClearFilesByExtension strSourcePath & "\exportspecs", "spec"
    
    ' Loop through specs
    For Each oSpec In CurrentProject.ImportExportSpecifications
    
        ' We may hit an error if the export does not have a description.
        ' Not sure how to test this other than a resume next.
        On Error Resume Next
        
        ' Build data to export
        With cData
            strXML = oSpec.XML
            .Add oSpec.Name
            .Add vbCrLf
            .Add Nz(oSpec.Description, vbNullString)
            .Add vbCrLf
            .Add vbCrLf
            .Add strXML
        End With
        
        ' Restore normal error handling
        If Err Then Err.Clear
        On Error GoTo 0
        
        ' Determine if this was an import or an export spec.
        If InStr(1, strXML, "</Export") = 0 Then
            strFolder = cOptions.GetExportFolder & "\importspecs"
        Else
            strFolder = cOptions.GetExportFolder & "\exportspecs"
        End If
        
        ' Build folder and file name
        VerifyPath strFolder
        strFile = GetSafeFileName(oSpec.Name) & ".spec"
        
        ' Save as file (including more than just xml)
        WriteFile cData.GetStr, strFolder & "\" & strFile
        cData.Clear
        Log.Add "  " & oSpec.Name, cOptions.ShowDebug
        
    Next oSpec
    
    If cOptions.ShowDebug Then
        Log.Add "[" & CurrentProject.ImportExportSpecifications.Count & "] specs exported."
    Else
        Log.Add "[" & CurrentProject.ImportExportSpecifications.Count & "]"
    End If
    
End Sub