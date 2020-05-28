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
                                       
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpWideCharStr As LongPtr, _
        ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As LongPtr, _
        ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As Long, _
        ByVal lpUsedDefaultChar As Long) As Long

    Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpMultiByteStr As LongPtr, _
        ByVal cchMultiByte As Long, _
        ByVal lpWideCharStr As LongPtr, _
        ByVal cchWideChar As Long _
        ) As Long
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

''' WinApi function that maps a UTF-16 (wide character) string to a new character string
    Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpWideCharStr As Long, _
        ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As Long, _
        ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As Long, _
        ByVal lpUsedDefaultChar As Long) As Long
#End If



' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001

' Cache the Ucs2 requirement for this database
Private m_blnUcs2 As Boolean
Private m_strDbPath As String


'---------------------------------------------------------------------------------------
' Procedure : RequiresUcs2
' Author    : Adam Waller
' Date      : 5/5/2020
' Purpose   : Returns true if the current database requires objects to be converted
'           : to Ucs2 format before importing. (Caching value for subsequent calls.)
'           : While this involves creating a new querydef object each time, the idea
'           : is that this would be faster than exporting a form if no queries exist
'           : in the current database.
'---------------------------------------------------------------------------------------
'
Public Function RequiresUcs2(Optional blnUseCache As Boolean = True) As Boolean

    Dim strTempFile As String
    Dim frm As Access.Form
    Dim strName As String
    Dim dbs As DAO.Database
    
    ' See if we already have a cached value
    If (m_strDbPath <> CurrentProject.FullName) Or Not blnUseCache Then
    
        ' Get temp file name
        strTempFile = GetTempFile
        
        ' Can't create querydef objects in ADP databases, so we have to use something else.
        If CurrentProject.ProjectType = acADP Then
            ' Create and export a blank form object.
            ' Turn of screen updates to improve performance and avoid flash.
            DoCmd.Echo False
            'strName = "frmTEMP_UCS2_" & Round(Timer)
            Set frm = Application.CreateForm
            strName = frm.Name
            DoCmd.Close acForm, strName, acSaveYes
            Application.SaveAsText acForm, strName, strTempFile
            DoCmd.DeleteObject acForm, strName
            DoCmd.Echo True
        Else
            ' Standard MDB database.
            ' Create and export a querydef object. Fast and light.
            strName = "qryTEMP_UCS2_" & Round(Timer)
            Set dbs = CurrentDb
            dbs.CreateQueryDef strName, "SELECT 1"
            Application.SaveAsText acQuery, strName, strTempFile
            dbs.QueryDefs.Delete strName
        End If
        
        ' Test and delete temp file
        m_strDbPath = CurrentProject.FullName
        m_blnUcs2 = FileIsUCS2Format(strTempFile)
        Kill strTempFile

    End If

    ' Return cached value
    RequiresUcs2 = m_blnUcs2
    
End Function


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

    Dim strText As String
    Dim utf8Bytes() As Byte
    Dim fnum As Integer
    
    ' Make sure the path exists before we write a file.
    VerifyPath FSO.GetParentFolderName(strDestinationFile)
    
    If FSO.FileExists(strDestinationFile) Then Kill strDestinationFile
    
    ' Check the first couple characters in the file for a UCS BOM.
    If FileIsUCS2Format(strSourceFile) Then
    
        ' Read file contents and delete (temp) source file
        With FSO.OpenTextFile(strSourceFile, , , TristateTrue)
            strText = .ReadAll
            .Close
        End With
        
        Kill strSourceFile
        
        utf8Bytes = Utf8BytesFromString(strText)
        
        
        ' Write as UTF-8
        fnum = FreeFile
                
        Open strDestinationFile For Binary As #fnum
        Put #fnum, 1, utf8Bytes
        Close fnum
        
    Else
        ' No conversion needed, move to destination.
        FSO.MoveFile strSourceFile, strDestinationFile
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

    Dim strText As String
    Dim utf8Bytes() As Byte
    Dim fnum As Integer

    ' Make sure the path exists before we write a file.
    VerifyPath FSO.GetParentFolderName(strDestinationFile)
    
    If FSO.FileExists(strDestinationFile) Then Kill strDestinationFile
    
    If FileIsUCS2Format(strSourceFile) Then
        ' No conversion needed, send to destination as is
        FSO.MoveFile strSourceFile, strDestinationFile
    Else
        
        ' Read file contents
        fnum = FreeFile
        
        Open strSourceFile For Binary As fnum
        ReDim utf8Bytes(LOF(fnum) - 1)
        Get fnum, , utf8Bytes
        Close fnum
                
        strText = Utf8BytesToString(utf8Bytes)
        
        ' Write as UCS-2 LE (BOM)
        With FSO.CreateTextFile(strDestinationFile, True, TristateTrue)
            .Write strText
            .Close
        End With
        
        Kill strSourceFile
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
    Dim UTF8BOM As String
    UTF8BOM = Chr(239) & Chr(187) & Chr(191) ' == &HEFBBBF
    Dim fileBOM As String
    fileBOM = Left$(fileContents, 3)
    
    If fileBOM = UTF8BOM Then
        RemoveUTF8BOM = Mid$(fileContents, 4)
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

'---------------------------------------------------------------------------------------
' Procedure : BytesLength
' Author    : Casper Englund
' Date      : 2020/05/01
' Purpose   : Return length of byte array
'---------------------------------------------------------------------------------------
Private Function BytesLength(abBytes() As Byte) As Long
    
    ' Trap error if array is uninitialized
    On Error Resume Next
    BytesLength = UBound(abBytes) - LBound(abBytes) + 1
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : Utf8BytesToString
' Author    : Adapted by Casper Englund
' Date      : 2020/05/01
' Purpose   : Return VBA "Unicode" string from byte array encoded in UTF-8
'---------------------------------------------------------------------------------------
Public Function Utf8BytesToString(abUtf8Array() As Byte) As String
    
    Dim nBytes As Long
    Dim nChars As Long
    Dim strOut As String
    Dim bUtf8Bom As Boolean
    
    Utf8BytesToString = ""
    
    ' Catch uninitialized input array
    nBytes = BytesLength(abUtf8Array)
    If nBytes <= 0 Then Exit Function
    bUtf8Bom = abUtf8Array(0) = 239 _
      And abUtf8Array(1) = 187 _
      And abUtf8Array(2) = 191
    
    If bUtf8Bom Then
        Dim i As Long
        Dim abTempArr() As Byte
        ReDim abTempArr(BytesLength(abUtf8Array) - 3)
        For i = 3 To UBound(abUtf8Array)
            abTempArr(i - 3) = abUtf8Array(i)
        Next i
        abUtf8Array = abTempArr
    End If
    
    ' Get number of characters in output string
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, 0&, 0&)
    
    ' Dimension output buffer to receive string
    strOut = String(nChars, 0)
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, StrPtr(strOut), nChars)
    Utf8BytesToString = Left$(strOut, nChars)

End Function

'---------------------------------------------------------------------------------------
' Procedure : Utf8BytesFromString
' Author    : Adapted by Casper Englund
' Date      : 2020/05/01
' Purpose   : Return byte array with VBA "Unicode" string encoded in UTF-8
'---------------------------------------------------------------------------------------
Public Function Utf8BytesFromString(strInput As String) As Byte()

    Dim nBytes As Long
    Dim abBuffer() As Byte
    
    ' Catch empty or null input string
    Utf8BytesFromString = vbNullString
    If Len(strInput) < 1 Then Exit Function
    
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, 0&, 0&, 0&, 0&)
    
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
    
End Function