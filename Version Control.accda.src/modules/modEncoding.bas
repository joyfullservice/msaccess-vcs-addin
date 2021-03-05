'---------------------------------------------------------------------------------------
' Module    : modEncoding
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Functions for reading and converting file encodings (Unicode, UTF-8)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit


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
                cData.Add .Read(clngChunkSize)  ' 128K
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
        ' Encode as UCS2-LE (UTF-16 LE)
        ReEncodeFile strSourceFile, "UTF-8", strDestinationFile, "UTF-16"
    
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
    
    ' Perform file conversion
    ReEncodeFile strSourceFile, "_autodetect_all", strDestinationFile, "UTF-8", adSaveCreateOverWrite

    ' Remove original file if specified.
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
    ReEncodeFile strSourceFile, "UTF-8", strDestinationFile, "_autodetect_all", adSaveCreateOverWrite
    
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
' Purpose   : Check for the specified BOM by reading the first few bytes in the file.
'---------------------------------------------------------------------------------------
'
Private Function FileHasBom(strFilePath As String, strBom As String) As Boolean
    FileHasBom = (strBom = StrConv(GetFileBytes(strFilePath, Len(strBom)), vbUnicode))
End Function


'---------------------------------------------------------------------------------------
' Procedure : StringHasExtendedASCII
' Author    : Adam Waller
' Date      : 3/6/2020
' Purpose   : Returns true if the string contains non-ASCI characters.
'---------------------------------------------------------------------------------------
'
Public Function StringHasExtendedASCII(strText As String) As Boolean

    Perf.OperationStart "Extended Chars Check"
    With New VBScript_RegExp_55.RegExp
        ' Include extended ASCII characters here.
        .Pattern = "[^\u0000-\u007F]"
        StringHasExtendedASCII = .Test(strText)
    End With
    Perf.OperationEnd
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReEncodeFile
' Author    : Adam Kauffman / Adam Waller
' Date      : 3/4/2021
' Purpose   : Change File Encoding. It reads and writes at the same time so the files must be different.
'---------------------------------------------------------------------------------------
'
Public Sub ReEncodeFile(strInputFile As String, strInputCharset As String, _
    strOutputFile As String, strOutputCharset As String, _
    Optional intOverwriteMode As SaveOptionsEnum = adSaveCreateOverWrite)

    Dim objOutputStream As ADODB.Stream
    
    ' Open streams and copy data
    Perf.OperationStart "Enc " & _
        Replace(strInputCharset, "_autodetect_all", "AUTO") & " as " & _
        Replace(strOutputCharset, "_autodetect_all", "AUTO")
    Set objOutputStream = New ADODB.Stream
    With New ADODB.Stream
        .Open
        .Type = adTypeBinary
        .LoadFromFile strInputFile
        .Type = adTypeText
        .Charset = strInputCharset
        objOutputStream.Open
        objOutputStream.Charset = strOutputCharset
        ' Copy data over by chunks to boost performance
        Do While .EOS <> True
            .CopyTo objOutputStream, clngChunkSize
        Loop
        .Close
    End With
    
    ' Save file and log performance
    objOutputStream.SaveToFile strOutputFile, intOverwriteMode
    objOutputStream.Close
    Perf.OperationEnd
    
End Sub