Option Compare Database
Option Private Module
Option Explicit


Public Sub ExportDataMacros(tableName As String, directory As String)
    On Error GoTo Err_export:
    
    Dim MacroConst As Integer
    Dim filePath As String: filePath = directory & tableName & ".xml"

    modFunctions.ExportObject MacroConst, tableName, filePath, modFileAccess.UsingUcs2
    FormatDataMacro filePath

    Exit Sub

Err_export:
    
End Sub


Public Sub ImportDataMacros(tableName As String, directory As String)
    On Error GoTo Err_import:
    Dim filePath As String: filePath = directory & tableName & ".xml"
    modFunctions.ImportObject MacroConst, tableName, filePath, modFileAccess.UsingUcs2

Err_import:
    
End Sub


'Splits exported DataMacro XML onto multiple lines
'Allows git to find changes within lines using diff
Private Sub FormatDataMacro(filePath As String)

    Dim saveStream As Object ' ADODB.Stream

    Set saveStream = CreateObject("ADODB.Stream")
    saveStream.Charset = "utf-8"
    saveStream.Type = 2 'adTypeText
    saveStream.Open

    Dim objStream As Object ' ADODB.Stream
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
