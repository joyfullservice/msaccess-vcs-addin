Option Explicit
Option Compare Database
Option Private Module


#If Mac Then
    ' Mac not supported
#ElseIf Win64 Then
    Private Declare PtrSafe _
        Function getTempPath Lib "kernel32" _
             Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                                   ByVal lpBuffer As String) As Long
    Private Declare PtrSafe _
        Function getTempFileName Lib "kernel32" _
             Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                                       ByVal lpPrefixString As String, _
                                       ByVal wUnique As Long, _
                                       ByVal lpTempFileName As String) As Long
#Else
    Private Declare _
        Function getTempPath Lib "kernel32" _
             Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                                   ByVal lpBuffer As String) As Long
    Private Declare _
        Function getTempFileName Lib "kernel32" _
             Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                                       ByVal lpPrefixString As String, _
                                       ByVal wUnique As Long, _
                                       ByVal lpTempFileName As String) As Long
#End If


' Determine if this database imports/exports code as UCS-2-LE. (Older file
' formats cause exported objects to use a Windows 8-bit character set.)
Public Function UsingUcs2(Optional ByRef appInstance As Application) As Boolean
    If appInstance Is Nothing Then Set appInstance = Application.Application
    
    Dim obj_name As String
    Dim obj_type As Variant
    Dim obj_type_split() As String
    Dim obj_type_name As String
    Dim obj_type_num As Long
    Dim thisDb As Database
    Set thisDb = appInstance.CurrentDb

    If CurrentProject.ProjectType = acMDB Then
        If thisDb.QueryDefs.Count > 0 Then
            obj_type_num = acQuery
            obj_name = thisDb.QueryDefs(0).Name
        Else
            For Each obj_type In Split( _
                "Forms|" & acForm & "," & _
                "Reports|" & acReport & "," & _
                "Scripts|" & acMacro & "," _
            )
                DoEvents
                obj_type_split = Split(obj_type, "|")
                obj_type_name = obj_type_split(0)
                obj_type_num = Val(obj_type_split(1))
                If thisDb.Containers(obj_type_name).Documents.Count > 0 Then
                    obj_name = thisDb.Containers(obj_type_name).Documents(0).Name
                    Exit For
                End If
            Next
        End If
    Else
        ' ADP Project
        If CurrentData.AllQueries.Count > 0 Then
            obj_type_num = acServerView
            obj_name = CurrentData.AllQueries(1).Name
        ElseIf CurrentProject.AllForms.Count > 0 Then
            ' Try a form
            obj_type_num = acForm
            obj_name = CurrentProject.AllForms(1).Name
        Else
            ' Can add more object types as needed...
        End If
    End If

    Dim tempFileName As String: tempFileName = GetTempFile()
    
    If obj_name = "" Then
        ' No objects found, make one to test.
        obj_name = "Temp_Test_Query_Delete_Me"
        
        thisDb.CreateQueryDef obj_name, "SELECT * FROM TEST WHERE TESTING=TRUE"
        appInstance.SaveAsText acQuery, obj_name, tempFileName
        thisDb.QueryDefs.Delete obj_name
    Else
        ' Use found object
        appInstance.SaveAsText obj_type_num, obj_name, tempFileName
    End If

    UsingUcs2 = FileIsUCS2Format(tempFileName)
    
    FSO.DeleteFile tempFileName
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : FileIsUCS2Format
' Author    : Adam Kauffman
' Date      : 02/24/2020
' Purpose   : Check the file header for USC-2-LE BOM marker and return true if found.
'---------------------------------------------------------------------------------------
'
Public Function FileIsUCS2Format(ByVal theFilePath As String) As Boolean
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Dim bytes As String
    bytes = "  "
    Open theFilePath For Binary Access Read As fileNumber
    Get fileNumber, 1, bytes
    Close fileNumber
    
    FileIsUCS2Format = (Asc(Mid(bytes, 1, 1)) = &HFF) And (Asc(Mid(bytes, 2, 1)) = &HFE)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ConvertUcs2Utf8
' Author    : Adam Waller
' Date      : 1/23/2019
' Purpose   : Convert a UCS2-little-endian encoded file to UTF-8.
'---------------------------------------------------------------------------------------
'
Public Sub ConvertUcs2Utf8(strSourceFile As String, strDestinationFile As String)
    If FileIsUCS2Format(strSourceFile) Then
        Dim stmNew As Object
        Set stmNew = CreateObject("ADODB.Stream")
        Dim strText As String
        
        ' Read file contents
        With FSO.OpenTextFile(strSourceFile, , , TristateTrue)
            strText = .ReadAll
            .Close
        End With
        
        ' Write as UTF-8
        With stmNew
            .Open
            .Type = 2 'adTypeText
            .Charset = "utf-8"
            .WriteText strText
            .SaveToFile strDestinationFile, 2 'adSaveCreateOverWrite
            .Close
        End With
        
        Set stmNew = Nothing
    Else
        ' No conversion needed, send to destination as is
        FSO.CopyFile strSourceFile, strDestinationFile
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ConvertUtf8Ucs2
' Author    : Adam Waller
' Date      : 1/24/2019
' Purpose   : Convert the file to old UCS-2 unicode format
'---------------------------------------------------------------------------------------
'
Public Sub ConvertUtf8Ucs2(strSourceFile As String, strDestinationFile As String)
    If FileIsUCS2Format(strSourceFile) Then
        ' No conversion needed, send to destination as is
        FSO.CopyFile strSourceFile, strDestinationFile
    Else
        Dim stmNew As Object
        Set stmNew = CreateObject("ADODB.Stream")
        Dim strText As String
        
        ' Read file contents
        With FSO.OpenTextFile(strSourceFile, , , TristateFalse)
            strText = RemoveUTF8BOM(.ReadAll)
            .Close
        End With
        
        ' Write as UCS-2 LE (BOM)
        With stmNew
            .Open
            .Type = 2 'adTypeText
            .Charset = "unicode"  ' The original Windows "Unicode" was UCS-2
            .WriteText strText
            .SaveToFile strDestinationFile, 2  'adSaveCreateOverWrite
            .Close
        End With
        
        Set stmNew = Nothing
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveUTF8BOM
' Author    : Adam Kauffman
' Date      : 1/24/2019
' Purpose   : Will remove a UTF8 BOM from the start of the string passed in.
'---------------------------------------------------------------------------------------
'
Public Function RemoveUTF8BOM(ByVal fileContents As String) As String
    Const UTF8BOM As String = "﻿"
    If Left$(fileContents, 3) = UTF8BOM Then
        RemoveUTF8BOM = Right$(fileContents, Len(fileContents) - 3)
    Else ' No BOM detected
        RemoveUTF8BOM = fileContents
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetTempFile
' Author    : Adapted by Adam Waller
' Date      : 1/23/2019
' Purpose   : Generate Random / Unique temporary file name.
'---------------------------------------------------------------------------------------
'
Public Function GetTempFile(Optional strPrefix As String = "VBA") As String

    Dim strPath As String * 512
    Dim strName As String * 576
    Dim lngReturn As Long
    
    lngReturn = getTempPath(512, strPath)
    lngReturn = getTempFileName(strPath, strPrefix, 0, strName)
    If lngReturn <> 0 Then GetTempFile = Left$(strName, InStr(strName, vbNullChar) - 1)
    
End Function