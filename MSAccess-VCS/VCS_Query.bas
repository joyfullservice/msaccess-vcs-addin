Attribute VB_Name = "VCS_Query"
Option Compare Database
Option Explicit

Public Sub ExportQueryAsSQL(qry As QueryDef, ByVal file_path As String, _
                            Optional ByVal Ucs2Convert As Boolean = False)

    VCS_Dir.VCS_MkDirIfNotExist Left$(file_path, InStrRev(file_path, "\"))
    If Ucs2Convert Then
        Dim tempFileName As String
        tempFileName = VCS_File.VCS_TempFile()
        writeTextToFile qry.sql, tempFileName
        VCS_File.VCS_ConvertUcs2Utf8 tempFileName, file_path
    Else
        writeTextToFile qry.sql, file_path
    End If

End Sub


Private Sub writeTextToFile(ByVal text As String, ByVal file_path As String)
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(file_path)

    oFile.WriteLine text
    oFile.Close
    
    Set fso = Nothing
    Set oFile = Nothing

End Sub

Private Function readFromTextFile(ByVal file_path As String) As String
    
    Dim textRead As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.OpenTextFile(file_path, ForReading)
    
    
    Do While Not oFile.AtEndOfStream

        textRead = textRead & oFile.ReadLine & vbCrLf
    
    Loop

    readFromTextFile = textRead
    
    oFile.Close
    
    Set fso = Nothing
    Set oFile = Nothing

End Function



Public Sub ImportQueryFromSQL(ByVal obj_name As String, ByVal file_path As String, _
                                Optional ByVal Ucs2Convert As Boolean = False)

    If Not VCS_Dir.VCS_FileExists(file_path) Then Exit Sub
    
    Dim qry As QueryDef
    
    If Ucs2Convert Then
        
        Dim tempFileName As String
        tempFileName = VCS_File.VCS_TempFile()
        VCS_File.VCS_ConvertUtf8Ucs2 file_path, tempFileName
        CurrentDb.CreateQueryDef obj_name, readFromTextFile(file_path)
        
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.DeleteFile tempFileName
    Else
        CurrentDb.CreateQueryDef obj_name, readFromTextFile(file_path)
    End If

End Sub

