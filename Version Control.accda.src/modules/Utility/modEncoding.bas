Attribute VB_Name = "modEncoding"
'---------------------------------------------------------------------------------------
' Module    : modEncoding
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Functions for reading and converting file encodings (Unicode, UTF-8)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Utility")


' API call to determine active code page (default system encoding)
Private Declare PtrSafe Function GetACP Lib "kernel32" () As Long


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

            ' Log performance
            Perf.OperationStart "Unicode Conversion"

            ' Read file contents and delete (temp) source file
            Set cData = New clsConcat
            With FSO.OpenTextFile(strSourceFile, ForReading, False, intTristate)
                ' Read chunks of text, rather than the whole thing at once for massive
                ' performance gains when reading large files.
                ' See https://docs.microsoft.com/is-is/sql/ado/reference/ado-api/readtext-method
                Do While Not .AtEndOfStream
                    cData.Add .Read(CHUNK_SIZE)  ' 128K
                Loop
                .Close
            End With

            ' Write as UTF-8 in the destination file.
            ' (Path will be verified before writing)
            WriteFile cData.GetStr, strDestinationFile
            Perf.OperationEnd

        Else
            ' Fully encoded as UTF-16
            ReEncodeFile strSourceFile, "utf-16", strDestinationFile, "utf-8"
        End If

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
        ReEncodeFile strSourceFile, "utf-8", strDestinationFile, "utf-16"

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
    ReEncodeFile strSourceFile, GetSystemEncoding, strDestinationFile, "utf-8", adSaveCreateOverWrite

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
    ReEncodeFile strSourceFile, "utf-8", strDestinationFile, GetSystemEncoding, adSaveCreateOverWrite

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
    Dim bte() As Byte
    bte = GetFileBytes(strFilePath, 3)
    HasUtf8Bom = ((bte(0) = &HEF) And (bte(1) = &HBB) And (bte(2) = &HBF))
End Function


'---------------------------------------------------------------------------------------
' Procedure : HasUcs2Bom
' Author    : Adam Waller
' Date      : 8/1/2020
' Purpose   : Returns true if the file begins with the bytes `FF FE` (ÿþ)
'           : Note that these must be read as bytes, not compared to a string value
'           : if the system is using the UTF-8 (beta) option in Windows 10. See #378
'---------------------------------------------------------------------------------------
'
Public Function HasUcs2Bom(strFilePath As String) As Boolean
    Dim bte() As Byte
    bte = GetFileBytes(strFilePath, 2)
    HasUcs2Bom = ((bte(0) = &HFF) And (bte(1) = &HFE))
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
    Perf.OperationStart "Enc. " & strInputCharset & " as " & strOutputCharset
    Set objOutputStream = New ADODB.Stream
    With New ADODB.Stream
        .Open
        .Type = adTypeBinary
        Perf.OperationStart "Load Stream from File"
        .LoadFromFile strInputFile
        Perf.OperationEnd
        .Type = adTypeText
        .Charset = strInputCharset
        objOutputStream.Open
        objOutputStream.Charset = strOutputCharset
        ' Copy from one stream to the other
        Perf.OperationStart "Copy Stream"
        .CopyTo objOutputStream
        Perf.OperationEnd
        .Close
    End With

    ' Save file and log performance
    VerifyPath strOutputFile
    Perf.OperationStart "Save Stream to File"
    objOutputStream.SaveToFile strOutputFile, intOverwriteMode
    Perf.OperationEnd
    objOutputStream.Close
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetSystemEncoding
' Author    : Adam Waller
' Date      : 7/1/2021
' Purpose   : Return the current encoding type used for non-UTF-8 text files.
'           : (Such as VBA code modules.)
'           : https://docs.microsoft.com/en-us/windows/win32/intl/code-page-identifiers
'           : https://documentation.help/MS-Office-VB/ofhowConstants.htm
'           : * Note that using utf-8 as a default system encoding may not work
'           : correctly with some extended characters in VBA code modules. The VBA IDE
'           : does not support Unicode characters, and requires code pages to display
'           : extended/non-English characters. See Issues #60, #186, #180, #246, #377
'---------------------------------------------------------------------------------------
'
Public Function GetSystemEncoding(Optional blnAllowUtf8 As Boolean = False) As String

    Static lngEncoding As Long

    ' Call API to determine active code page, caching return value.
    If lngEncoding = 0 Then lngEncoding = GetACP
    Select Case lngEncoding

        ' Language encoding mappings are defined here, based on the following sources:
        ' https://docs.microsoft.com/en-us/office/vba/api/office.msoencoding
        ' https://docs.microsoft.com/en-us/dotnet/api/system.text.encoding?view=net-5.0
        Case msoEncodingEBCDICUSCanada:                 GetSystemEncoding = "IBM037"
        Case msoEncodingOEMUnitedStates:                GetSystemEncoding = "IBM437"
        Case msoEncodingEBCDICInternational:            GetSystemEncoding = "IBM500"
        Case msoEncodingArabicASMO:                     GetSystemEncoding = "ASMO-708"
        Case msoEncodingArabicTransparentASMO:          GetSystemEncoding = "DOS-720"
        Case msoEncodingOEMGreek437G:                   GetSystemEncoding = "ibm737"
        Case msoEncodingOEMBaltic:                      GetSystemEncoding = "ibm775"
        Case msoEncodingOEMMultilingualLatinI:          GetSystemEncoding = "ibm850"
        Case msoEncodingOEMMultilingualLatinII:         GetSystemEncoding = "ibm852"
        Case msoEncodingOEMCyrillic:                    GetSystemEncoding = "IBM855"
        Case msoEncodingOEMTurkish:                     GetSystemEncoding = "ibm857"
        Case msoEncodingOEMPortuguese:                  GetSystemEncoding = "IBM860"
        Case msoEncodingOEMIcelandic:                   GetSystemEncoding = "ibm861"
        Case msoEncodingOEMHebrew:                      GetSystemEncoding = "DOS-862"
        Case msoEncodingOEMCanadianFrench:              GetSystemEncoding = "IBM863"
        Case msoEncodingOEMArabic:                      GetSystemEncoding = "IBM864"
        Case msoEncodingOEMNordic:                      GetSystemEncoding = "IBM865"
        Case msoEncodingOEMCyrillicII:                  GetSystemEncoding = "cp866"
        Case msoEncodingOEMModernGreek:                 GetSystemEncoding = "ibm869"
        Case msoEncodingEBCDICMultilingualROECELatin2:  GetSystemEncoding = "IBM870"
        Case msoEncodingThai:                           GetSystemEncoding = "windows-874"
        Case msoEncodingEBCDICGreekModern:              GetSystemEncoding = "cp875"
        Case msoEncodingJapaneseShiftJIS:               GetSystemEncoding = "shift_jis"
        Case msoEncodingSimplifiedChineseGBK:           GetSystemEncoding = "gb2312"
        Case msoEncodingKorean:                         GetSystemEncoding = "ks_c_5601-1987"
        Case msoEncodingTraditionalChineseBig5:         GetSystemEncoding = "big5"
        Case msoEncodingEBCDICTurkishLatin5:            GetSystemEncoding = "IBM1026"
        Case msoEncodingUnicodeLittleEndian:            GetSystemEncoding = "utf-16"
        Case msoEncodingUnicodeBigEndian:               GetSystemEncoding = "unicodeFFFE"
        Case msoEncodingCentralEuropean:                GetSystemEncoding = "windows-1250"
        Case msoEncodingCyrillic:                       GetSystemEncoding = "windows-1251"
        Case msoEncodingWestern:                        GetSystemEncoding = "Windows-1252"
        Case msoEncodingGreek:                          GetSystemEncoding = "windows-1253"
        Case msoEncodingTurkish:                        GetSystemEncoding = "windows-1254"
        Case msoEncodingHebrew:                         GetSystemEncoding = "windows-1255"
        Case msoEncodingArabic:                         GetSystemEncoding = "windows-1256"
        Case msoEncodingBaltic:                         GetSystemEncoding = "windows-1257"
        Case msoEncodingVietnamese:                     GetSystemEncoding = "windows-1258"
        Case msoEncodingKoreanJohab:                    GetSystemEncoding = "Johab"
        Case msoEncodingMacRoman:                       GetSystemEncoding = "macintosh"
        Case msoEncodingMacJapanese:                    GetSystemEncoding = "x-mac-japanese"
        Case msoEncodingMacTraditionalChineseBig5:      GetSystemEncoding = "x-mac-chinesetrad"
        Case msoEncodingMacKorean:                      GetSystemEncoding = "x-mac-korean"
        Case msoEncodingMacArabic:                      GetSystemEncoding = "x-mac-arabic"
        Case msoEncodingMacHebrew:                      GetSystemEncoding = "x-mac-hebrew"
        Case msoEncodingMacGreek1:                      GetSystemEncoding = "x-mac-greek"
        Case msoEncodingMacCyrillic:                    GetSystemEncoding = "x-mac-cyrillic"
        Case msoEncodingMacSimplifiedChineseGB2312:     GetSystemEncoding = "x-mac-chinesesimp"
        Case msoEncodingMacRomania:                     GetSystemEncoding = "x-mac-romanian"
        Case msoEncodingMacUkraine:                     GetSystemEncoding = "x-mac-ukrainian"
        Case msoEncodingMacLatin2:                      GetSystemEncoding = "x-mac-ce"
        Case msoEncodingMacIcelandic:                   GetSystemEncoding = "x-mac-icelandic"
        Case msoEncodingMacTurkish:                     GetSystemEncoding = "x-mac-turkish"
        Case msoEncodingMacCroatia:                     GetSystemEncoding = "x-mac-croatian"
        Case msoEncodingTaiwanCNS:                      GetSystemEncoding = "x-Chinese-CNS"
        Case msoEncodingTaiwanTCA:                      GetSystemEncoding = "x-cp20001"
        Case msoEncodingTaiwanEten:                     GetSystemEncoding = "x-Chinese-Eten"
        Case msoEncodingTaiwanIBM5550:                  GetSystemEncoding = "x-cp20003"
        Case msoEncodingTaiwanTeleText:                 GetSystemEncoding = "x-cp20004"
        Case msoEncodingTaiwanWang:                     GetSystemEncoding = "x-cp20005"
        Case msoEncodingIA5IRV:                         GetSystemEncoding = "x-IA5"
        Case msoEncodingIA5German:                      GetSystemEncoding = "x-IA5-German"
        Case msoEncodingIA5Swedish:                     GetSystemEncoding = "x-IA5-Swedish"
        Case msoEncodingIA5Norwegian:                   GetSystemEncoding = "x-IA5-Norwegian"
        Case msoEncodingUSASCII:                        GetSystemEncoding = "us-ascii"
        Case msoEncodingT61:                            GetSystemEncoding = "x-cp20261"
        Case msoEncodingISO6937NonSpacingAccent:        GetSystemEncoding = "x-cp20269"
        Case msoEncodingEBCDICGermany:                  GetSystemEncoding = "IBM273"
        Case msoEncodingEBCDICDenmarkNorway:            GetSystemEncoding = "IBM277"
        Case msoEncodingEBCDICFinlandSweden:            GetSystemEncoding = "IBM278"
        Case msoEncodingEBCDICItaly:                    GetSystemEncoding = "IBM280"
        Case msoEncodingEBCDICLatinAmericaSpain:        GetSystemEncoding = "IBM284"
        Case msoEncodingEBCDICUnitedKingdom:            GetSystemEncoding = "IBM285"
        Case msoEncodingEBCDICJapaneseKatakanaExtended: GetSystemEncoding = "IBM290"
        Case msoEncodingEBCDICFrance:                   GetSystemEncoding = "IBM297"
        Case msoEncodingEBCDICArabic:                   GetSystemEncoding = "IBM420"
        Case msoEncodingEBCDICGreek:                    GetSystemEncoding = "IBM423"
        Case msoEncodingEBCDICHebrew:                   GetSystemEncoding = "IBM424"
        Case msoEncodingEBCDICKoreanExtended:           GetSystemEncoding = "x-EBCDIC-KoreanExtended"
        Case msoEncodingEBCDICThai:                     GetSystemEncoding = "IBM-Thai"
        Case msoEncodingKOI8R:                          GetSystemEncoding = "koi8-r"
        Case msoEncodingEBCDICIcelandic:                GetSystemEncoding = "IBM871"
        Case msoEncodingEBCDICRussian:                  GetSystemEncoding = "IBM880"
        Case msoEncodingEBCDICTurkish:                  GetSystemEncoding = "IBM905"
        Case msoEncodingEBCDICSerbianBulgarian:         GetSystemEncoding = "cp1025"
        Case msoEncodingKOI8U:                          GetSystemEncoding = "koi8-u"
        Case msoEncodingISO88591Latin1:                 GetSystemEncoding = "iso-8859-1"
        Case msoEncodingISO88592CentralEurope:          GetSystemEncoding = "iso-8859-2"
        Case msoEncodingISO88593Latin3:                 GetSystemEncoding = "iso-8859-3"
        Case msoEncodingISO88594Baltic:                 GetSystemEncoding = "iso-8859-4"
        Case msoEncodingISO88595Cyrillic:               GetSystemEncoding = "iso-8859-5"
        Case msoEncodingISO88596Arabic:                 GetSystemEncoding = "iso-8859-6"
        Case msoEncodingISO88597Greek:                  GetSystemEncoding = "iso-8859-7"
        Case msoEncodingISO88598Hebrew:                 GetSystemEncoding = "iso-8859-8"
        Case msoEncodingISO88599Turkish:                GetSystemEncoding = "iso-8859-9"
        Case msoEncodingISO885915Latin9:                GetSystemEncoding = "iso-8859-15"
        Case msoEncodingEuropa3:                        GetSystemEncoding = "x-Europa"
        Case msoEncodingISO88598HebrewLogical:          GetSystemEncoding = "iso-8859-8-i"
        Case msoEncodingISO2022JPNoHalfwidthKatakana:   GetSystemEncoding = "iso-2022-jp"
        Case msoEncodingISO2022JPJISX02021984:          GetSystemEncoding = "csISO2022JP"
        Case msoEncodingISO2022JPJISX02011989:          GetSystemEncoding = "iso-2022-jp"
        Case msoEncodingISO2022KR:                      GetSystemEncoding = "iso-2022-kr"
        Case msoEncodingISO2022CNTraditionalChinese:    GetSystemEncoding = "x-cp50227"
        Case msoEncodingEUCJapanese:                    GetSystemEncoding = "euc-jp"
        Case msoEncodingEUCChineseSimplifiedChinese:    GetSystemEncoding = "EUC-CN"
        Case msoEncodingEUCKorean:                      GetSystemEncoding = "euc-kr"
        Case msoEncodingHZGBSimplifiedChinese:          GetSystemEncoding = "hz-gb-2312"
        Case msoEncodingSimplifiedChineseGB18030:       GetSystemEncoding = "GB18030"
        Case msoEncodingISCIIDevanagari:                GetSystemEncoding = "x-iscii-de"
        Case msoEncodingISCIIBengali:                   GetSystemEncoding = "x-iscii-be"
        Case msoEncodingISCIITamil:                     GetSystemEncoding = "x-iscii-ta"
        Case msoEncodingISCIITelugu:                    GetSystemEncoding = "x-iscii-te"
        Case msoEncodingISCIIAssamese:                  GetSystemEncoding = "x-iscii-as"
        Case msoEncodingISCIIOriya:                     GetSystemEncoding = "x-iscii-or"
        Case msoEncodingISCIIKannada:                   GetSystemEncoding = "x-iscii-ka"
        Case msoEncodingISCIIMalayalam:                 GetSystemEncoding = "x-iscii-ma"
        Case msoEncodingISCIIGujarati:                  GetSystemEncoding = "x-iscii-gu"
        Case msoEncodingISCIIPunjabi:                   GetSystemEncoding = "x-iscii-pa"
        Case msoEncodingUTF7:                           GetSystemEncoding = "utf-7"

        ' In Windows 10, this is shown as a checkbox in Region settings for
        ' "Beta: Use Unicode UTF-8 for worldwide language support"
        Case msoEncodingUTF8:
            If blnAllowUtf8 Then
                GetSystemEncoding = "utf-8"
            Else
                ' If UTF-8 is not allowed (such as for code modules), then fall back
                ' to most commonly used codepage, supporting most Western Euorpean
                ' languages, but not Cyrillic. https://www.wikiwand.com/en/Windows-1252
                GetSystemEncoding = "Windows-1252"
            End If

        ' Any other language encoding not defined above (should be very rare)
        Case Else
            ' Attempt to autodetect the language based on the content.
            ' (Note that this does not work as well on code as it does
            '  with normal written language. See issue #186)
            GetSystemEncoding = "_autodetect_all"
    End Select

End Function


'---------------------------------------------------------------------------------------
' Function  : EscapeXmlName
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Escape a field or table name the way Access does for XML data export
'           : element names (_xHHHH_ for disallowed characters).
'---------------------------------------------------------------------------------------
'
Public Function EscapeXmlName(strName As String) As String

    Const CHAR_UNDERSCORE As Long = 95
    Const CHAR_LOWER_X As Long = 120

    Dim lngPos As Long
    Dim lngChar As Long
    Dim blnEscape As Boolean
    Dim cOut As New clsConcat

    For lngPos = 1 To Len(strName)
        lngChar = AscW(Mid$(strName, lngPos, 1))
        If lngChar = CHAR_UNDERSCORE Then
            ' Access escapes an underscore only when a lowercase "x" follows it, so the
            ' pair cannot be misread as the start of an _xHHHH_ escape sequence. An
            ' uppercase "X" is left alone, so this test must be case sensitive.
            blnEscape = False
            If lngPos < Len(strName) Then
                blnEscape = (AscW(Mid$(strName, lngPos + 1, 1)) = CHAR_LOWER_X)
            End If
        Else
            blnEscape = Not IsXmlExportNameChar(lngChar)
        End If
        If blnEscape Then
            cOut.Add "_x", Right$("0000" & Hex$(lngChar And &HFFFF&), 4), "_"
        Else
            cOut.Add ChrW$(lngChar)
        End If
    Next lngPos

    EscapeXmlName = cOut.GetStr

End Function


'---------------------------------------------------------------------------------------
' Function  : UnescapeXmlName
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Reverse of EscapeXmlName. Converts _xHHHH_ sequences in an exported XML
'           : element or index-key name back to the original character.
'           :
'           : A sequence is only decoded when it is well formed (_x, four hex digits, and
'           : a closing underscore). Anything else is literal text, which is what makes
'           : EscapeXmlName's rule -- escape an underscore only when a lowercase "x"
'           : follows it -- round-trip: "a_x1" escapes to "a_x005F_x1", and the "_x1"
'           : tail here is left alone because "x1" is not four hex digits.
'           :
'           : Comparisons are binary because Access escapes "_x" but not "_X", and this
'           : module compiles with Option Compare Database.
'---------------------------------------------------------------------------------------
'
Public Function UnescapeXmlName(strName As String) As String

    Const ESCAPE_LEN As Long = 7    ' _xHHHH_

    Dim lngPos As Long
    Dim lngLen As Long
    Dim lngCode As Long
    Dim cOut As New clsConcat

    lngLen = Len(strName)
    lngPos = 1

    Do While lngPos <= lngLen
        lngCode = -1
        If lngPos + ESCAPE_LEN - 1 <= lngLen Then
            If StrComp(Mid$(strName, lngPos, 2), "_x", vbBinaryCompare) = 0 _
                And Mid$(strName, lngPos + 6, 1) = "_" Then
                lngCode = HexWordValue(Mid$(strName, lngPos + 2, 4))
            End If
        End If
        If lngCode >= 0 Then
            cOut.Add ChrW$(lngCode)
            lngPos = lngPos + ESCAPE_LEN
        Else
            cOut.Add Mid$(strName, lngPos, 1)
            lngPos = lngPos + 1
        End If
    Loop

    UnescapeXmlName = cOut.GetStr

End Function


'---------------------------------------------------------------------------------------
' Function  : HexWordValue
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Return the value of exactly four hexadecimal digits, or -1 when the string
'           : is not four hex digits. Returning the value and the verdict together is the
'           : point: a caller cannot validate one string and convert a different one.
'           :
'           : The digits are accumulated arithmetically rather than passed to CLng with an
'           : "&H" prefix. CLng does handle that prefix correctly over the full 0000-FFFF
'           : range, but it raises a type mismatch on a trailing "&" -- the type-suffix
'           : character is VBA source-literal syntax, not part of string-to-number
'           : conversion. Doing the arithmetic keeps this off VBA's parser entirely.
'---------------------------------------------------------------------------------------
'
Private Function HexWordValue(strValue As String) As Long

    Dim lngPos As Long
    Dim lngChar As Long
    Dim lngDigit As Long
    Dim lngResult As Long

    HexWordValue = -1
    If Len(strValue) <> 4 Then Exit Function

    For lngPos = 1 To 4
        lngChar = AscW(Mid$(strValue, lngPos, 1))
        Select Case lngChar
            Case 48 To 57:  lngDigit = lngChar - 48     ' 0-9
            Case 65 To 70:  lngDigit = lngChar - 55     ' A-F
            Case 97 To 102: lngDigit = lngChar - 87     ' a-f
            Case Else:      Exit Function
        End Select
        lngResult = (lngResult * 16) + lngDigit
    Next lngPos

    HexWordValue = lngResult

End Function


'---------------------------------------------------------------------------------------
' Function  : NormalizeXmlSortValue
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Normalize a table-data XML element value for deterministic lexical sorting.
'           : Null/missing fields pass blnIsNull=True.
'---------------------------------------------------------------------------------------
'
Public Function NormalizeXmlSortValue(strValue As String, intFieldType As Integer, Optional blnIsNull As Boolean = False) As String

    Const MAX_SORT_LEN As Long = 255

    If blnIsNull Then
        NormalizeXmlSortValue = Chr$(1) & "N"
        Exit Function
    End If

    Select Case intFieldType
        Case dbBoolean
            If StrComp(strValue, "true", vbTextCompare) = 0 _
                Or strValue = "1" _
                Or StrComp(strValue, "-1", vbTextCompare) = 0 Then
                NormalizeXmlSortValue = Chr$(2) & "1"
            Else
                NormalizeXmlSortValue = Chr$(2) & "0"
            End If
        Case dbByte, dbInteger, dbLong, dbSingle, dbDouble, dbCurrency, dbDecimal, dbNumeric
            NormalizeXmlSortValue = Chr$(3) & NormalizeNumericXmlSortValue(strValue)
        Case dbDate
            NormalizeXmlSortValue = Chr$(4) & Left$(strValue, MAX_SORT_LEN)
        Case Else
            NormalizeXmlSortValue = Chr$(5) & Left$(strValue, MAX_SORT_LEN)
    End Select

End Function


'---------------------------------------------------------------------------------------
' Function  : ComposeXmlSortKey
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Concatenate normalized sort-field parts into one lexical key. Each part
'           : is terminated with vbNullChar so multi-field keys stay unambiguous without
'           : making string length the primary sort dimension.
'---------------------------------------------------------------------------------------
'
Public Function ComposeXmlSortKey(colParts As Collection, lngOrdinal As Long) As String

    Const ORDINAL_WIDTH As Long = 10

    Dim varPart As Variant
    Dim cKey As New clsConcat

    For Each varPart In colParts
        cKey.Add CStr(varPart), vbNullChar
    Next varPart

    ComposeXmlSortKey = cKey.GetStr & Format$(lngOrdinal, String$(ORDINAL_WIDTH, "0"))

End Function


'---------------------------------------------------------------------------------------
' Function  : XmlSortKeyOrdinal
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Recover the stable-sort ordinal suffix from a ComposeXmlSortKey result.
'---------------------------------------------------------------------------------------
'
Public Function XmlSortKeyOrdinal(strKey As String) As Long

    Const ORDINAL_WIDTH As Long = 10

    XmlSortKeyOrdinal = CLng(Right$(strKey, ORDINAL_WIDTH))

End Function


'---------------------------------------------------------------------------------------
' Function  : NormalizeNumericXmlSortValue
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Fixed-width numeric normalization so lexical order matches numeric order.
'---------------------------------------------------------------------------------------
'
Public Function NormalizeNumericXmlSortValue(strValue As String) As String

    ' Fractional digits are zero-padded rather than optional (########) so that every key
    ' is the same length. Without padding, a nines-complemented "2" ("...97") would sort
    ' ahead of "2.5" ("...97.4"), reversing the intended order for negative values.
    Const NUM_SORT_FORMAT As String = "0000000000000000.00000000"

    Dim dblVal As Double

    If Len(strValue) = 0 Then
        NormalizeNumericXmlSortValue = "P" & Format$(0, NUM_SORT_FORMAT)
        Exit Function
    End If

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    dblVal = CDbl(strValue)
    If Err Then
        Err.Clear
        NormalizeNumericXmlSortValue = "T" & Left$(strValue, 255)
        Exit Function
    End If

    If dblVal < 0 Then
        ' Nines-complement the magnitude so that larger negatives sort first. Subtracting
        ' the value from a large constant cannot be used here, because Format$ rounds a
        ' Double to 15 significant digits and the 16-digit results collapse together.
        NormalizeNumericXmlSortValue = "N" & NinesComplement(Format$(-dblVal, NUM_SORT_FORMAT))
    Else
        NormalizeNumericXmlSortValue = "P" & Format$(dblVal, NUM_SORT_FORMAT)
    End If

End Function


'---------------------------------------------------------------------------------------
' Function  : NinesComplement
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Replace each digit with (9 - digit), leaving the decimal separator alone,
'           : so that a descending magnitude produces an ascending lexical key.
'---------------------------------------------------------------------------------------
'
Private Function NinesComplement(strNumber As String) As String

    Dim lngPos As Long
    Dim lngChar As Long
    Dim cOut As New clsConcat

    For lngPos = 1 To Len(strNumber)
        lngChar = AscW(Mid$(strNumber, lngPos, 1))
        If lngChar >= 48 And lngChar <= 57 Then
            cOut.Add ChrW$(105 - lngChar)
        Else
            cOut.Add ChrW$(lngChar)
        End If
    Next lngPos

    NinesComplement = cOut.GetStr

End Function


'---------------------------------------------------------------------------------------
' Function  : IsXmlExportNameChar
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Characters Access leaves unescaped in exported XML element names.
'---------------------------------------------------------------------------------------
'
Private Function IsXmlExportNameChar(lngChar As Long) As Boolean
    IsXmlExportNameChar = _
        (lngChar >= 48 And lngChar <= 57) _
        Or (lngChar >= 65 And lngChar <= 90) _
        Or (lngChar >= 97 And lngChar <= 122) _
        Or lngChar = 95
End Function
