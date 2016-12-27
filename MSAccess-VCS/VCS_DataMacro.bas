Attribute VB_Name = "VCS_DataMacro"
Option Compare Database

Option Private Module
Option Explicit


' For Access 2007 (VBA6) and earlier
#If Not VBA7 Then
  Private Const acTableDataMacro As Integer = 12
#End If

Public Sub VCS_ExportDataMacros(ByVal tableName As String, ByVal directory As String)
    On Error GoTo Err_export
    Dim filePath As String

    filePath = directory & tableName & ".dm"

    VCS_IE_Functions.VCS_ExportObject acTableDataMacro, tableName, filePath, VCS_File.VCS_UsingUcs2
    FormatDataMacro filePath

    Exit Sub

Err_export:
    ' Error to export dataMacro, no contains dataMacro. Do nothing
End Sub

Public Sub VCS_ImportDataMacros(ByVal tableName As String, ByVal directory As String)
    On Error GoTo Err_import
    Dim filePath As String

    filePath = directory & tableName & ".dm"
    VCS_IE_Functions.VCS_ImportObject acTableDataMacro, tableName, filePath, VCS_File.VCS_UsingUcs2
    
    Exit Sub
    
Err_import:
    ' Error to import dataMacro. Do nothing
End Sub

'Splits exported DataMacro XML onto multiple lines
'Allows git to find changes within lines using diff
Private Sub FormatDataMacro(ByVal filePath As String)

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
            If tag <> vbNullString Then
                saveStream.WriteText tag & ">", 1 'adWriteLine
            End If
        Next
        
    Loop
    
    objStream.Close
    saveStream.SaveToFile filePath, 2 'adSaveCreateOverWrite
    saveStream.Close

End Sub