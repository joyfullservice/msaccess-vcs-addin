'---------------------------------------------------------------------------------------
' Module    : modEncoding
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Functions for reading and converting file encodings (Unicode, UTF-8)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


''' Maps a character string to a UTF-16 (wide character) string
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long _
    ) As Long

''' WinApi function that maps a UTF-16 (wide character) string to a new character string
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long _
    ) As Long


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
            Perf.OperationStart "App.SaveAsText()"
            Application.SaveAsText acForm, strName, strTempFile
            Perf.OperationEnd
            DoCmd.DeleteObject acForm, strName
            DoCmd.Echo True
        Else
            ' Standard MDB database.
            ' Create and export a querydef object. Fast and light.
            strName = "qryTEMP_UCS2_" & Round(Timer)
            Set dbs = CurrentDb
            dbs.CreateQueryDef strName, "SELECT 1"
            Perf.OperationStart "App.SaveAsText()"
            Application.SaveAsText acQuery, strName, strTempFile
            Perf.OperationEnd
            dbs.QueryDefs.Delete strName
        End If
        
        ' Test and delete temp file
        m_strDbPath = CurrentProject.FullName
        m_blnUcs2 = HasUcs2Bom(strTempFile)
        DeleteFile strTempFile, True

    End If

    ' Return cached value
    RequiresUcs2 = m_blnUcs2
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : ConvertUcs2Utf8
' Author    : Adam Waller
' Date      : 1/23/2019
' Purpose   : Convert a UCS2-little-endian encoded file to UTF-8.
'           : Typically the source file will be a temp file.
'---------------------------------------------------------------------------------------
'
Public Sub ConvertUcs2Utf8(strSourceFile As String, strDestinationFile As String, _
    Optional blnDeleteSourceFileAfterConversion As Boolean = True)

    Dim cData As clsConcat
    Dim blnIsAdp As Boolean
    Dim intTristate As Tristate
    
    ' Remove any existing file.
    If FSO.FileExists(strDestinationFile) Then DeleteFile strDestinationFile, True
    
    ' ADP Projects do not use the UCS BOM, but may contain mixed UTF-16 content
    ' representing unicode characters.
    blnIsAdp = (CurrentProject.ProjectType = acADP)
    
    ' Check the first couple characters in the file for a UCS BOM.
    If HasUcs2Bom(strSourceFile) Or blnIsAdp Then
    
        ' Determine format
        If blnIsAdp Then
            ' Possible mixed UTF-16 content
            intTristate = TristateMixed
        Else
            ' Fully encoded as UTF-16
            intTristate = TristateTrue
        End If
        
        ' Log performance
        Perf.OperationStart "Unicode Conversion"
        
        ' Read file contents and delete (temp) source file
        Set cData = New clsConcat
        With FSO.OpenTextFile(strSourceFile, ForReading, False, intTristate)
            ' Read chunks of text, rather than the whole thing at once for massive
            ' performance gains when reading large files.
            ' See https://docs.microsoft.com/is-is/sql/ado/reference/ado-api/readtext-method
            Do While Not .AtEndOfStream
                cData.Add .Read(131072) ' 128K
            Loop
            .Close
        End With
        
        ' Write as UTF-8 in the destination file.
        ' (Path will be verified before writing)
        WriteFile cData.GetStr, strDestinationFile
        Perf.OperationEnd
        
        ' Remove the source (temp) file if specified
        If blnDeleteSourceFileAfterConversion Then DeleteFile strSourceFile, True
    Else
        ' No conversion needed, move/copy to destination.
        VerifyPath strDestinationFile
        If blnDeleteSourceFileAfterConversion Then
            FSO.MoveFile strSourceFile, strDestinationFile
        Else
            FSO.CopyFile strSourceFile, strDestinationFile
        End If
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ConvertUtf8Ucs2
' Author    : Adam Waller
' Date      : 1/24/2019
' Purpose   : Convert the file to old UCS-2 unicode format.
'           : Typically the destination file will be a temp file.
'---------------------------------------------------------------------------------------
'
Public Sub ConvertUtf8Ucs2(strSourceFile As String, strDestinationFile As String, _
    Optional blnDeleteSourceFileAfterConversion As Boolean = True)

    Dim strText As String
    Dim utf8Bytes() As Byte

    ' Make sure the path exists before we write a file.
    VerifyPath strDestinationFile
    If FSO.FileExists(strDestinationFile) Then DeleteFile strDestinationFile, True
    
    If HasUcs2Bom(strSourceFile) Then
        ' No conversion needed, move/copy to destination.
        If blnDeleteSourceFileAfterConversion Then
            FSO.MoveFile strSourceFile, strDestinationFile
        Else
            FSO.CopyFile strSourceFile, strDestinationFile
        End If
    Else
        ' Monitor performance
        Perf.OperationStart "Unicode Conversion"
        
        ' Read file contents and convert byte array to string
        utf8Bytes = GetFileBytes(strSourceFile)
        strText = Utf8BytesToString(utf8Bytes)
        
        ' Write as UCS-2 LE (BOM)
        With FSO.CreateTextFile(strDestinationFile, True, TristateTrue)
            .Write strText
            .Close
        End With
        Perf.OperationEnd
        
        ' Remove original file if specified.
        If blnDeleteSourceFileAfterConversion Then DeleteFile strSourceFile, True
    End If
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ConvertAnsiiUtf8
' Author    : Adam Waller
' Date      : 2/3/2021
' Purpose   : Convert an ANSI encoded file to UTF-8. This allows extended characters
'           : to properly display in diffs and other programs. See issue #154
'---------------------------------------------------------------------------------------
'
Public Sub ConvertAnsiUtf8(strSourceFile As String, strDestinationFile As String, _
    Optional blnDeleteSourceFileAfterConversion As Boolean = True)
    
    ' Convert the ANSI content to UTF-8, and write to a new file.
    ' (Adds UTF-8 BOM if extended characters are used.)
    WriteFile ReadFile(strSourceFile, "_autodetect_all"), strDestinationFile
    If blnDeleteSourceFileAfterConversion Then DeleteFile strSourceFile
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ConvertUtf8Ansii
' Author    : Adam Waller
' Date      : 2/3/2021
' Purpose   : Convert a UTF-8 file back to ANSI.
'---------------------------------------------------------------------------------------
'
Public Sub ConvertUtf8Ansi(strSourceFile As String, strDestinationFile As String, _
    Optional blnDeleteSourceFileAfterConversion As Boolean = True)
    
    ' Perform file conversion
    Perf.OperationStart "ANSI Conversion"
    With New ADODB.Stream
        .Charset = "_autodetect_all"
        .Open
        .WriteText ReadFile(strSourceFile)
        .SaveToFile strDestinationFile, adSaveCreateOverWrite
        .Close
    End With
    Perf.OperationEnd
    
    ' Remove original file if specified.
    If blnDeleteSourceFileAfterConversion Then DeleteFile strSourceFile
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : HasUtf8Bom
' Author    : Adam Waller
' Date      : 7/30/2020
' Purpose   : Returns true if the file begins with a UTF-8 BOM
'---------------------------------------------------------------------------------------
'
Public Function HasUtf8Bom(strFilePath As String) As Boolean
    HasUtf8Bom = FileHasBom(strFilePath, UTF8_BOM)
End Function


'---------------------------------------------------------------------------------------
' Procedure : HasUcs2Bom
' Author    : Adam Waller
' Date      : 8/1/2020
' Purpose   : Returns true if the file begins with
'---------------------------------------------------------------------------------------
'
Public Function HasUcs2Bom(strFilePath As String) As Boolean
    HasUcs2Bom = FileHasBom(strFilePath, UCS2_BOM)
End Function


'---------------------------------------------------------------------------------------
' Procedure : FileHasBom
' Author    : Adam Waller
' Date      : 8/1/2020
' Purpose   : Check for the specified BOM
'---------------------------------------------------------------------------------------
'
Private Function FileHasBom(strFilePath As String, strBom As String) As Boolean
    Dim strFound As String
    strFound = StrConv((GetFileBytes(strFilePath, Len(strBom))), vbUnicode)
    FileHasBom = (strFound = strBom)
End Function


'---------------------------------------------------------------------------------------
' Procedure : RemoveUTF8BOM
' Author    : Adam Kauffman
' Date      : 1/24/2019
' Purpose   : Will remove a UTF8 BOM from the start of the string passed in.
'---------------------------------------------------------------------------------------
'
Public Function RemoveUTF8BOM(ByVal fileContents As String) As String
    Dim UTF8BOM As String
    UTF8BOM = Chr$(239) & Chr$(187) & Chr$(191) ' == &HEFBBBF
    Dim fileBOM As String
    fileBOM = Left$(fileContents, 3)
    
    If fileBOM = UTF8BOM Then
        RemoveUTF8BOM = Mid$(fileContents, 4)
    Else ' No BOM detected
        RemoveUTF8BOM = fileContents
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : BytesLength
' Author    : Casper Englund
' Date      : 2020/05/01
' Purpose   : Return length of byte array
'---------------------------------------------------------------------------------------
Private Function BytesLength(abBytes() As Byte) As Long
    
    ' Ignore error if array is uninitialized
    On Error Resume Next
    BytesLength = UBound(abBytes) - LBound(abBytes) + 1
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    
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
    
    Utf8BytesToString = vbNullString
    
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
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(strInput), -1, 0&, 0&, 0&, 0&)
    
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadFile
' Author    : Adam Waller / Indigo
' Date      : 11/4/2020
' Purpose   : Read text file.
'           : Read in UTF-8 encoding, removing a BOM if found at start of file.
'---------------------------------------------------------------------------------------
'
Public Function ReadFile(strPath As String, Optional strCharset As String = "UTF-8") As String

    Dim stm As ADODB.Stream
    Dim strText As String
    Dim cData As clsConcat
    Dim strBom As String
    
    ' Get BOM header, if applicable
    Select Case strCharset
        Case "UTF-8": strBom = UTF8_BOM
        Case "Unicode": strBom = UCS2_BOM
    End Select
    
    If FSO.FileExists(strPath) Then
        Perf.OperationStart "Read File"
        Set cData = New clsConcat
        Set stm = New ADODB.Stream
        With stm
            .Charset = strCharset
            .Open
            .LoadFromFile strPath
            ' Check for BOM
            If strBom <> vbNullString Then
                strText = .ReadText(Len(strBom))
                If strText <> strBom Then cData.Add strText
            End If
            ' Read chunks of text, rather than the whole thing at once for massive
            ' performance gains when reading large files.
            ' See https://docs.microsoft.com/is-is/sql/ado/reference/ado-api/readtext-method
            Do While Not .EOS
                cData.Add .ReadText(131072) ' 128K
            Loop
            .Close
        End With
        Set stm = Nothing
        Perf.OperationEnd
    End If
    
    ' Return text contents of file.
    ReadFile = cData.GetStr
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : WriteFile
' Author    : Adam Waller
' Date      : 1/23/2019
' Purpose   : Save string variable to text file. (Building the folder path if needed)
'           : Saves in UTF-8 encoding, adding a BOM if extended or unicode content
'           : is found in the file. https://stackoverflow.com/a/53036838/4121863
'---------------------------------------------------------------------------------------
'
Public Sub WriteFile(strText As String, strPath As String)

    Dim strContent As String
    Dim bteUtf8() As Byte
    
    ' Ensure that we are ending the content with a vbcrlf
    strContent = strText
    If Right$(strText, 2) <> vbCrLf Then strContent = strContent & vbCrLf

    ' Build a byte array from the text
    bteUtf8 = Utf8BytesFromString(strContent)
    
    ' Write binary content to file.
    WriteBinaryFile bteUtf8, StringHasUnicode(strContent), strPath
        
End Sub


'---------------------------------------------------------------------------------------
' Procedure : WriteBinaryFile
' Author    : Adam Waller
' Date      : 8/3/2020
' Purpose   : Write binary content to a file with optional UTF-8 BOM.
'---------------------------------------------------------------------------------------
'
Public Sub WriteBinaryFile(bteContent() As Byte, blnUtf8Bom As Boolean, strPath As String)

    Dim stm As ADODB.Stream
    Dim bteBOM(0 To 2) As Byte
    
    ' Write to a binary file using a Stream object
    Set stm = New ADODB.Stream
    With stm
        .Type = adTypeBinary
        .Open
        If blnUtf8Bom Then
            bteBOM(0) = &HEF
            bteBOM(1) = &HBB
            bteBOM(2) = &HBF
            .Write bteBOM
        End If
        .Write bteContent
        VerifyPath strPath
        Perf.OperationStart "Write to Disk"
        .SaveToFile strPath, adSaveCreateOverWrite
        Perf.OperationEnd
    End With
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : StringHasUnicode
' Author    : Adam Waller
' Date      : 3/6/2020
' Purpose   : Returns true if the string contains non-ASCI characters.
'---------------------------------------------------------------------------------------
'
Public Function StringHasUnicode(strText As String) As Boolean
    
    Dim reg As VBScript_RegExp_55.RegExp
    
    Perf.OperationStart "Unicode Check"
    Set reg = New VBScript_RegExp_55.RegExp
    With reg
        ' Include extended ASCII characters here.
        .Pattern = "[^\u0000-\u007F]"
        StringHasUnicode = .Test(strText)
    End With
    Perf.OperationEnd
    
End Function