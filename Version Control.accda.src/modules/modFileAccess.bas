Option Compare Database
Option Private Module
Option Explicit

#If Mac Then
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

' --------------------------------
' Structures
' --------------------------------

' Structure to track buffered reading or writing of binary files
Private Type BinFile
    file_num As Integer
    file_len As Long
    file_pos As Long
    buffer As String
    buffer_len As Integer
    buffer_pos As Integer
    at_eof As Boolean
    mode As String
End Type

' --------------------------------
' Basic functions missing from VB 6: buffered file read/write, string builder, encoding check & conversion
' --------------------------------

' Open a binary file for reading (mode = 'r') or writing (mode = 'w').
Private Function BinOpen(file_path As String, mode As String) As BinFile
    Dim f As BinFile

    f.file_num = FreeFile
    f.mode = LCase(mode)
    If f.mode = "r" Then
        Open file_path For Binary Access Read As f.file_num
        f.file_len = LOF(f.file_num)
        f.file_pos = 0
        If f.file_len > &H4000 Then
            f.buffer = String(&H4000, " ")
            f.buffer_len = &H4000
        Else
            f.buffer = String(f.file_len, " ")
            f.buffer_len = f.file_len
        End If
        f.buffer_pos = 0
        Get f.file_num, f.file_pos + 1, f.buffer
    Else
        DelIfExist file_path
        Open file_path For Binary Access Write As f.file_num
        f.file_len = 0
        f.file_pos = 0
        f.buffer = String(&H4000, " ")
        f.buffer_len = 0
        f.buffer_pos = 0
    End If

    BinOpen = f
End Function

' Buffered read one byte at a time from a binary file.
Private Function BinRead(ByRef f As BinFile) As Integer
    If f.at_eof = True Then
        BinRead = 0
        Exit Function
    End If

    BinRead = Asc(Mid(f.buffer, f.buffer_pos + 1, 1))

    f.buffer_pos = f.buffer_pos + 1
    If f.buffer_pos >= f.buffer_len Then
        f.file_pos = f.file_pos + &H4000
        If f.file_pos >= f.file_len Then
            f.at_eof = True
            Exit Function
        End If
        If f.file_len - f.file_pos > &H4000 Then
            f.buffer_len = &H4000
        Else
            f.buffer_len = f.file_len - f.file_pos
            f.buffer = String(f.buffer_len, " ")
        End If
        f.buffer_pos = 0
        Get f.file_num, f.file_pos + 1, f.buffer
    End If
End Function

' Buffered write one byte at a time from a binary file.
Private Sub BinWrite(ByRef f As BinFile, b As Integer)
    Mid(f.buffer, f.buffer_pos + 1, 1) = Chr(b)
    f.buffer_pos = f.buffer_pos + 1
    If f.buffer_pos >= &H4000 Then
        Put f.file_num, , f.buffer
        f.buffer_pos = 0
    End If
End Sub

' Close binary file.
Private Sub BinClose(ByRef f As BinFile)
    If f.mode = "w" And f.buffer_pos > 0 Then
        f.buffer = Left(f.buffer, f.buffer_pos)
        Put f.file_num, , f.buffer
    End If
    Close f.file_num
End Sub


' Binary convert a UCS2-little-endian encoded file to UTF-8.
Public Sub ConvertUcs2Utf8(Source As String, dest As String)
    Dim f_in As BinFile, f_out As BinFile
    Dim in_low As Integer, in_high As Integer

    f_in = BinOpen(Source, "r")
    f_out = BinOpen(dest, "w")

    Do While Not f_in.at_eof
        in_low = BinRead(f_in)
        in_high = BinRead(f_in)
        If in_high = 0 And in_low < &H80 Then
            ' U+0000 - U+007F   0LLLLLLL
            BinWrite f_out, in_low
        ElseIf in_high < &H8 Then
            ' U+0080 - U+07FF   110HHHLL 10LLLLLL
            BinWrite f_out, &HC0 + ((in_high And &H7) * &H4) + ((in_low And &HC0) / &H40)
            BinWrite f_out, &H80 + (in_low And &H3F)
        Else
            ' U+0800 - U+FFFF   1110HHHH 10HHHHLL 10LLLLLL
            BinWrite f_out, &HE0 + ((in_high And &HF0) / &H10)
            BinWrite f_out, &H80 + ((in_high And &HF) * &H4) + ((in_low And &HC0) / &H40)
            BinWrite f_out, &H80 + (in_low And &H3F)
        End If
    Loop

    BinClose f_in
    BinClose f_out
End Sub

' Binary convert a UTF-8 encoded file to UCS2-little-endian.
Public Sub ConvertUtf8Ucs2(Source As String, dest As String)
    Dim f_in As BinFile, f_out As BinFile
    Dim in_1 As Integer, in_2 As Integer, in_3 As Integer

    f_in = BinOpen(Source, "r")
    f_out = BinOpen(dest, "w")

    Do While Not f_in.at_eof
        in_1 = BinRead(f_in)
        If (in_1 And &H80) = 0 Then
            ' U+0000 - U+007F   0LLLLLLL
            BinWrite f_out, in_1
            BinWrite f_out, 0
        ElseIf (in_1 And &HE0) = &HC0 Then
            ' U+0080 - U+07FF   110HHHLL 10LLLLLL
            in_2 = BinRead(f_in)
            BinWrite f_out, ((in_1 And &H3) * &H40) + (in_2 And &H3F)
            BinWrite f_out, (in_1 And &H1C) / &H4
        Else
            ' U+0800 - U+FFFF   1110HHHH 10HHHHLL 10LLLLLL
            in_2 = BinRead(f_in)
            in_3 = BinRead(f_in)
            BinWrite f_out, ((in_2 And &H3) * &H40) + (in_3 And &H3F)
            BinWrite f_out, ((in_1 And &HF) * &H10) + ((in_2 And &H3C) / &H4)
        End If
    Loop

    BinClose f_in
    BinClose f_out
End Sub



' Determine if this database imports/exports code as UCS-2-LE. (Older file
' formats cause exported objects to use a Windows 8-bit character set.)
Public Function UsingUcs2() As Boolean
    Dim obj_name As String, obj_type As Variant, fn As Integer, bytes As String
    Dim obj_type_split() As String, obj_type_name As String, obj_type_num As Integer

    If CurrentDb.QueryDefs.Count > 0 Then
        obj_type_num = acQuery
        obj_name = CurrentDb.QueryDefs(0).Name
    Else
        For Each obj_type In Split( _
            "Forms|" & acForm & "," & _
            "Reports|" & acReport & "," & _
            "Scripts|" & acMacro & "," & _
            "Modules|" & acModule _
        )
            DoEvents
            obj_type_split = Split(obj_type, "|")
            obj_type_name = obj_type_split(0)
            obj_type_num = Val(obj_type_split(1))
            If CurrentDb.Containers(obj_type_name).Documents.Count > 0 Then
                obj_name = CurrentDb.Containers(obj_type_name).Documents(0).Name
                Exit For
            End If
        Next
    End If

    If obj_name = "" Then
        ' No objects found that can be used to test UCS2 versus UTF-8
        UsingUcs2 = True
        Exit Function
    End If

    Dim tempFileName As String: tempFileName = TempFile()
    Application.SaveAsText obj_type_num, obj_name, tempFileName
    fn = FreeFile
    Open tempFileName For Binary Access Read As fn
    bytes = "  "
    Get fn, 1, bytes
    If Asc(Mid(bytes, 1, 1)) = &HFF And Asc(Mid(bytes, 2, 1)) = &HFE Then
        UsingUcs2 = True
    Else
        UsingUcs2 = False
    End If
    Close fn
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.DeleteFile (tempFileName)
End Function


' Generate Random / Unique tempprary file name.
Public Function TempFile(Optional sPrefix As String = "VBA") As String
Dim sTmpPath As String * 512
Dim sTmpName As String * 576
Dim nRet As Long
Dim sFileName As String
    
    nRet = getTempPath(512, sTmpPath)
    nRet = getTempFileName(sTmpPath, sPrefix, 0, sTmpName)
    If nRet <> 0 Then sFileName = Left$(sTmpName, InStr(sTmpName, vbNullChar) - 1)
    TempFile = sFileName
End Function