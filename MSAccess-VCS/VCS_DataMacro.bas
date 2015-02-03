Attribute VB_Name = "VCS_DataMacro"
Option Compare Database

Option Explicit


Public Sub ExportDataMacros(tableName As String, directory As String)
    On Error GoTo Err_export:

    Dim filePath As String: filePath = directory & tableName & ".xml"

    VCS_IE_Functions.ExportObject acTableDataMacro, tableName, filePath, VCS_File.UsingUcs2
    FormatDataMacro filePath

    Exit Sub

Err_export:
    
    
End Sub

Public Sub ImportDataMacros(tableName As String, directory As String)
    On Error GoTo Err_import:
    Dim filePath As String: filePath = directory & tableName & ".xml"
    VCS_IE_Functions.ImportObject acTableDataMacro, tableName, filePath, VCS_File.UsingUcs2

Err_import:
    
End Sub

'Splits exported DataMacro XML onto multiple lines
'Allows git to find changes within lines using diff
Private Sub FormatDataMacro(filePath As String)

    Dim saveStream As Object 'ADODB.Stream

    Set saveStream = CreateObject("ADODB.Stream")
    saveStream.Charset = "utf-8"
    saveStream.Type = 2 'adTypeText
    saveStream.Open

    Dim objStream As Object 'ADODB.Stream
    Dim strData As String
    Set objStream = CreateObject("ADODB.Stream")

    objStream.Charset = "utf-8"
    objStream.Type = 2 'adTypeText
    objStream.Open
    objStream.LoadFromFile (filePath)
    
    Do While Not objStream.EOS
        strData = objStream.ReadText(-2) 'adReadLine

        Dim tag As Variant
        
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












