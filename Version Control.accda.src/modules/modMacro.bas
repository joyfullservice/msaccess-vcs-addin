Option Explicit
Option Compare Database
Option Private Module


Public Sub ExportDataMacros(TableName As String, directory As String, cModel As IVersionControl, Optional Ucs2Convert As Boolean = False)
    'On Error GoTo Err_export:
    Dim blnSkip As Boolean
    
    Dim filePath As String: filePath = directory & TableName & ".xml"

    ' Check for fast save
    If Not cModel Is Nothing Then
        If cModel.FastSave Then
            blnSkip = (HasMoreRecentChanges(CurrentData.AllTables(TableName), filePath))
        End If
    End If
    
    If Not blnSkip Then
        modFunctions.ExportObject acTableDataMacro, TableName, filePath, cModel, Ucs2Convert
        If FileExists(filePath) Then FormatDataMacro filePath
    End If
    
    Exit Sub

Err_export:
    
End Sub


Public Sub ImportDataMacros(TableName As String, directory As String)
    On Error GoTo Err_import:
    Dim filePath As String: filePath = directory & TableName & ".xml"
    modFunctions.ImportObject MacroConst, TableName, filePath, modFileAccess.UsingUcs2

Err_import:
    
End Sub


'Splits exported DataMacro XML onto multiple lines
'Allows git to find changes within lines using diff
Private Sub FormatDataMacro(filePath As String)

    Dim saveStream As New ADODB.Stream
    Dim objStream As New ADODB.Stream
    Dim strData As String
    Dim tag As Variant
    
    saveStream.Charset = "utf-8"
    saveStream.Type = adTypeText
    saveStream.Open

    objStream.Charset = "utf-8"
    objStream.Type = adTypeText
    objStream.Open
    objStream.LoadFromFile filePath
    
    Do While Not objStream.EOS
        strData = objStream.ReadText(-2) 'adReadLine

        For Each tag In Split(strData, ">")
            If tag <> "" Then
                saveStream.WriteText tag & ">", 1 'adWriteLine
            End If
        Next
    Loop
    
    objStream.Close
    saveStream.SaveToFile filePath, 2 'adSaveCreateOverWrite
    saveStream.Close

End Sub


'---------------------------------------------------------------------------------------
' Procedure : MacroConst
' Author    : Adam Waller
' Date      : 5/14/2015
' Purpose   : Return version specific macro constant for export
'---------------------------------------------------------------------------------------
'
Private Function MacroConst() As Integer
    If Application.Version > 12 Then
        MacroConst = 12 ' acTableDataMacro
    Else
        ' Access 2007 and earlier
        MacroConst = acMacro
    End If
End Function