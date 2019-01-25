Option Explicit
Option Compare Database
Option Private Module


'---------------------------------------------------------------------------------------
' Procedure : ExportDataMacros
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub ExportDataMacros(strTable As String, strFolder As String, cModel As IVersionControl)
    
    Dim blnSkip As Boolean
    Dim strPath As String
    
    ' Build file path
    strPath = strFolder & GetSafeFileName(strTable) & ".xml"

    ' Check for fast save
    If cModel.FastSave Then blnSkip = Not (HasMoreRecentChanges(CurrentData.AllTables(strTable), strPath))
    
    If Not blnSkip Then
        ' Need to test whether to use Ucs2Convert on data macros.
        ExportObject acTableDataMacro, strTable, strPath, cModel
        If FSO.FileExists(strPath) Then FormatDataMacro strPath
    End If
        
End Sub


Public Sub ImportDataMacros(TableName As String, directory As String)
    Dim filePath As String: filePath = directory & TableName & ".xml"
    ImportObject acTableDataMacro, TableName, filePath, modFileAccess.UsingUcs2
End Sub



'---------------------------------------------------------------------------------------
' Procedure : FormatDataMacro
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Splits exported DataMacro XML onto multiple lines
'           : Allows git to find changes within lines using diff
'---------------------------------------------------------------------------------------
'
Private Sub FormatDataMacro(strFile As String)

    Dim stmRead As New ADODB.Stream
    Dim strData As String
    Dim varTag As Variant
    Dim cData As New clsConcat
    
    With stmRead
        .Charset = "utf-8"
        .Type = adTypeText
        .Open
        .LoadFromFile strFile
    End With
    
    Do While Not stmRead.EOS
        strData = stmRead.ReadText(adReadLine)
        For Each varTag In Split(strData, ">")
            If varTag <> "" Then
                cData.Add CStr(varTag)
                cData.Add ">"
                cData.Add vbCrLf
            End If
        Next
    Loop
    
    ' Close and delete file
    stmRead.Close
    Set stmRead = Nothing
    Kill strFile
    
    ' Write new file with formatted content
    WriteFile cData.GetStr, strFile
    
End Sub