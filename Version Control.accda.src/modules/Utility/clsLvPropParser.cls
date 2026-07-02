VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLvPropParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsLvPropParser
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Parses the undocumented LvProp binary blob from MSysObjects to extract
'           : locally-stored field and table display properties for linked tables.
'           : Uses cursor-based pure VBA byte math -- no API calls, no CopyMemory.
'           :
'           : The LvProp blob stores properties such as ColumnWidth, DisplayControl,
'           : RowSource, Description, Caption, lookup column definitions, datasheet
'           : formatting, and user-defined custom properties. For ODBC-linked tables,
'           : ALL display/lookup properties are stored here. For Access-linked tables,
'           : only locally-overridden properties are stored (layout, custom, etc.);
'           : lookup/formatting properties are inherited from the back-end.
'           :
'           : Binary format overview:
'           :   [4: "MR2\0" magic] [4: dict_block_size]
'           :   [2: marker 0x80 0x00] [dict entries: 2-byte len + UTF-16LE name ...]
'           :   [data blocks: field blocks (flag 0x0001) and table block (flag 0x0000)]
'           :   Each block: [4: size] [2: flag] [4: hdr_size] [2: name_len] [name] [entries]
'           :   Each entry: [2: size] [1: flags] [1: type] [2: index] [2: vallen] [value]
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Utility")

Private m_bteData() As Byte
Private m_lngPos As Long
Private m_lngSize As Long
Private m_asPropNames() As String
Private m_lngPropCount As Long
Private m_intSanitizeLevel As eSanitizeLevel


'---------------------------------------------------------------------------------------
' Procedure : ParseLvProp
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Main entry point. Parses an LvProp binary blob and returns a Dictionary:
'           :   "FieldProperties" -> Dictionary: fieldName -> Dict: propName -> {Type, Value}
'           :   "TableProperties" -> Dictionary: propName -> {Type, Value}
'---------------------------------------------------------------------------------------
'
Public Function ParseLvProp(bteData() As Byte, intSanitizeLevel As eSanitizeLevel) As Dictionary

    Dim dResult As New Dictionary
    Dim dFieldProps As New Dictionary
    Dim dTableProps As New Dictionary
    Dim lngDictBlockSize As Long
    Dim lngDataStart As Long

    m_intSanitizeLevel = intSanitizeLevel
    m_bteData = bteData
    m_lngSize = UBound(m_bteData) - LBound(m_bteData) + 1
    m_lngPos = LBound(m_bteData)

    ' Validate MR2 magic header
    If m_lngSize < 8 Then GoTo Done
    If m_bteData(m_lngPos) <> &H4D Or m_bteData(m_lngPos + 1) <> &H52 Or _
       m_bteData(m_lngPos + 2) <> &H32 Or m_bteData(m_lngPos + 3) <> 0 Then GoTo Done

    ' Read dict block size (includes these 4 bytes)
    m_lngPos = m_lngPos + 4
    lngDictBlockSize = ReadLong()
    If lngDictBlockSize < 6 Then GoTo Done

    ' Parse the property name dictionary
    lngDataStart = LBound(m_bteData) + 4 + lngDictBlockSize
    ParseDictionary lngDataStart

    ' Walk data blocks
    m_lngPos = lngDataStart
    ParseDataBlocks dFieldProps, dTableProps

    ' Build result
    If dFieldProps.Count > 0 Then dResult.Add "FieldProperties", dFieldProps
    If dTableProps.Count > 0 Then dResult.Add "TableProperties", dTableProps

Done:
    Set ParseLvProp = dResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : ParseDictionary
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Reads the property name dictionary from the blob. Each entry is a
'           : 2-byte UTF-16LE length followed by the name string. Names are stored
'           : in a 0-based array used to resolve property indices in data entries.
'---------------------------------------------------------------------------------------
'
Private Sub ParseDictionary(lngDataStart As Long)

    Dim lngLen As Long
    Dim lngIndex As Long
    Dim lngBase As Long

    lngBase = LBound(m_bteData)

    ' Skip the 2-byte marker (0x80, 0x00)
    m_lngPos = lngBase + 8 + 2

    ' Pre-size the name array (generous initial size, will ReDim later)
    ReDim m_asPropNames(0 To 127)
    lngIndex = 0

    Do While m_lngPos + 1 < lngDataStart
        lngLen = ReadInt()
        If lngLen <= 0 Or m_lngPos + lngLen > lngDataStart Then Exit Do
        If lngIndex > UBound(m_asPropNames) Then
            ReDim Preserve m_asPropNames(0 To lngIndex + 64)
        End If
        m_asPropNames(lngIndex) = ReadUtf16(lngLen)
        lngIndex = lngIndex + 1
    Loop

    m_lngPropCount = lngIndex

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ParseDataBlocks
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Walks the data section, parsing field blocks (flag=0x0001) and
'           : non-field/table blocks (flag=0x0000).
'---------------------------------------------------------------------------------------
'
Private Sub ParseDataBlocks(dFieldProps As Dictionary, dTableProps As Dictionary)

    Dim lngBlockSize As Long
    Dim lngFlag As Long
    Dim lngBlockEnd As Long
    Dim lngHeaderSize As Long
    Dim lngNameLen As Long
    Dim strFieldName As String
    Dim dProps As Dictionary

    Do While m_lngPos + 5 < m_lngSize + LBound(m_bteData)
        ' Read block header
        lngBlockSize = ReadLong()
        If lngBlockSize < 6 Or m_lngPos - 4 + lngBlockSize > m_lngSize + LBound(m_bteData) Then Exit Do
        lngBlockEnd = m_lngPos - 4 + lngBlockSize
        lngFlag = ReadInt()

        ' Read inner header (present in both field and non-field blocks)
        lngHeaderSize = ReadLong()
        lngNameLen = ReadInt()

        If lngFlag = 1 And lngNameLen > 0 Then
            ' Field block
            strFieldName = ReadUtf16(lngNameLen)
            Set dProps = ParsePropertyEntries(lngBlockEnd)
            If dProps.Count > 0 Then
                dFieldProps.Add strFieldName, SortDictionaryByKeys(dProps)
            End If
        ElseIf lngFlag = 0 Then
            ' Non-field block (table properties + NameMap)
            ParseTableEntries lngBlockEnd, dTableProps
        End If

        ' Advance to end of block
        m_lngPos = lngBlockEnd
    Loop

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ParsePropertyEntries
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Parses property entries within a field block, applying default-value
'           : filtering to keep the output concise.
'---------------------------------------------------------------------------------------
'
Private Function ParsePropertyEntries(lngBlockEnd As Long) As Dictionary

    Dim dProps As New Dictionary
    Dim dProp As Dictionary
    Dim lngEntrySize As Long
    Dim lngEntryEnd As Long
    Dim bteFlags As Byte
    Dim bteType As Byte
    Dim lngIndex As Long
    Dim lngValLen As Long
    Dim strName As String
    Dim varValue As Variant

    Do While m_lngPos + 7 < lngBlockEnd
        lngEntrySize = ReadInt()
        If lngEntrySize < 8 Or m_lngPos - 2 + lngEntrySize > lngBlockEnd Then Exit Do
        lngEntryEnd = m_lngPos - 2 + lngEntrySize

        bteFlags = ReadByte()
        bteType = ReadByte()
        lngIndex = ReadInt()
        lngValLen = ReadInt()

        ' Resolve property name from dictionary
        If lngIndex >= 0 And lngIndex < m_lngPropCount Then
            strName = m_asPropNames(lngIndex)
        Else
            strName = vbNullString
        End If

        ' Read value based on type
        varValue = ReadValue(bteType, lngValLen)

        ' Apply filtering: skip GUID, NameMap, and known default values
        If Len(strName) > 0 Then
            If Not ShouldSkipFieldProperty(strName, bteType, varValue) Then
                Set dProp = New Dictionary
                dProp.Add "Type", CLng(bteType)
                dProp.Add "Value", varValue
                dProps.Add strName, dProp
            End If
        End If

        m_lngPos = lngEntryEnd
    Loop

    Set ParsePropertyEntries = dProps

End Function


'---------------------------------------------------------------------------------------
' Procedure : ParseTableEntries
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Parses property entries from a non-field (table) block, skipping
'           : internal entries like NameMap and GUID.
'---------------------------------------------------------------------------------------
'
Private Sub ParseTableEntries(lngBlockEnd As Long, dTableProps As Dictionary)

    Dim dProp As Dictionary
    Dim lngEntrySize As Long
    Dim lngEntryEnd As Long
    Dim bteFlags As Byte
    Dim bteType As Byte
    Dim lngIndex As Long
    Dim lngValLen As Long
    Dim strName As String
    Dim varValue As Variant

    Do While m_lngPos + 7 < lngBlockEnd
        lngEntrySize = ReadInt()
        If lngEntrySize < 8 Or m_lngPos - 2 + lngEntrySize > lngBlockEnd Then Exit Do
        lngEntryEnd = m_lngPos - 2 + lngEntrySize

        bteFlags = ReadByte()
        bteType = ReadByte()
        lngIndex = ReadInt()
        lngValLen = ReadInt()

        ' Resolve property name
        If lngIndex >= 0 And lngIndex < m_lngPropCount Then
            strName = m_asPropNames(lngIndex)
        Else
            strName = vbNullString
        End If

        ' Read value
        varValue = ReadValue(bteType, lngValLen)

        ' Apply filtering
        If Len(strName) > 0 Then
            If Not ShouldSkipTableProperty(strName, bteType, varValue) Then
                If Not dTableProps.Exists(strName) Then
                    Set dProp = New Dictionary
                    dProp.Add "Type", CLng(bteType)
                    dProp.Add "Value", varValue
                    dTableProps.Add strName, dProp
                End If
            End If
        End If

        m_lngPos = lngEntryEnd
    Loop

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReadValue
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Reads a property value from the blob based on its DAO type code.
'---------------------------------------------------------------------------------------
'
Private Function ReadValue(bteType As Byte, lngValLen As Long) As Variant

    Dim lngVal As Long

    If lngValLen <= 0 Or m_lngPos + lngValLen > m_lngSize + LBound(m_bteData) Then
        ReadValue = Null
        Exit Function
    End If

    Select Case bteType
        Case 1  ' Boolean
            ReadValue = CBool(ReadByte() <> 0)
            m_lngPos = m_lngPos + lngValLen - 1

        Case 2  ' Integer
            If lngValLen >= 2 Then
                ReadValue = ReadInt()
                m_lngPos = m_lngPos + lngValLen - 2
            Else
                ReadValue = CInt(ReadByte())
                m_lngPos = m_lngPos + lngValLen - 1
            End If

        Case 3  ' Long
            If lngValLen >= 4 Then
                ReadValue = ReadLong()
                m_lngPos = m_lngPos + lngValLen - 4
            ElseIf lngValLen >= 2 Then
                ReadValue = CLng(ReadInt())
                m_lngPos = m_lngPos + lngValLen - 2
            Else
                ReadValue = CLng(ReadByte())
                m_lngPos = m_lngPos + lngValLen - 1
            End If

        Case 4  ' Long
            If lngValLen >= 4 Then
                ReadValue = ReadLong()
                m_lngPos = m_lngPos + lngValLen - 4
            Else
                ReadValue = CLng(ReadByte())
                m_lngPos = m_lngPos + lngValLen - 1
            End If

        Case 6  ' Single (IEEE 754 float)
            If lngValLen >= 4 Then
                lngVal = ReadLong()
                ReadValue = LongToSingle(lngVal)
                m_lngPos = m_lngPos + lngValLen - 4
            Else
                ReadValue = 0!
                m_lngPos = m_lngPos + lngValLen
            End If

        Case 10, 12 ' Text (UTF-16LE)
            ReadValue = ReadUtf16(lngValLen)

        Case Else
            ' Binary (type 9=GUID, 11=NameMap, etc.) -- skip
            ReadValue = Null
            m_lngPos = m_lngPos + lngValLen
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : ShouldSkipFieldProperty
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Returns True for field properties that should be excluded from export.
'           : Skips binary blobs (GUID), engine-internal properties, and properties
'           : at their known default values to keep JSON concise.
'---------------------------------------------------------------------------------------
'
Private Function ShouldSkipFieldProperty(strName As String, bteType As Byte, varValue As Variant) As Boolean

    ' Skip binary/internal types (GUID, NameMap)
    If bteType = 9 Or bteType = 11 Then
        ShouldSkipFieldProperty = True
        Exit Function
    End If

    If IsNull(varValue) Then
        ShouldSkipFieldProperty = True
        Exit Function
    End If

    ' Skip known engine-internal properties
    Select Case strName
        Case "Required", "AllowZeroLength", "Expression"
            ShouldSkipFieldProperty = True
            Exit Function
    End Select

    ' Skip properties at their default values
    Select Case strName
        Case "AggregateType":   ShouldSkipFieldProperty = (varValue = -1)
        Case "ResultType":      ShouldSkipFieldProperty = (varValue = 0)
        Case "CurrencyLCID":    ShouldSkipFieldProperty = (varValue = 0)
        Case "ColumnHidden":    ShouldSkipFieldProperty = IsDefaultBool(varValue, False)
        Case "ColumnWidth":     ShouldSkipFieldProperty = (varValue = -1)
        Case "ColumnOrder":     ShouldSkipFieldProperty = (varValue = 0)
        Case "TextAlign":       ShouldSkipFieldProperty = (varValue = 0)
        Case "DisplayControl":  ShouldSkipFieldProperty = (varValue = 109) ' acTextBox
        Case "ShowDatePicker":  ShouldSkipFieldProperty = (varValue = 1)
        Case "IMEMode":         ShouldSkipFieldProperty = (varValue = 0)
        Case "IMESentenceMode": ShouldSkipFieldProperty = (varValue = 3)
        Case "Caption", "Description"
            ShouldSkipFieldProperty = (Len(Nz(varValue)) = 0)
    End Select

    ' At Standard sanitize level or above, also skip properties that Access
    ' auto-determines when linking, based on the back-end column type.
    ' These are never manually customized and create significant JSON noise.
    ' If a future need arises to preserve manual overrides, remove these filters.
    If Not ShouldSkipFieldProperty And m_intSanitizeLevel >= eslStandard Then
        Select Case strName
            Case "UnicodeCompression": ShouldSkipFieldProperty = True
            Case "AppendOnly":  ShouldSkipFieldProperty = IsDefaultBool(varValue, False)
            Case "TextFormat":  ShouldSkipFieldProperty = (varValue = 0)
        End Select
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : ShouldSkipTableProperty
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Returns True for table-level properties that should be excluded.
'           : Skips binary blobs, internal metadata, and default values.
'---------------------------------------------------------------------------------------
'
Private Function ShouldSkipTableProperty(strName As String, bteType As Byte, varValue As Variant) As Boolean

    ' Skip binary/internal types (GUID, NameMap)
    If bteType = 9 Or bteType = 11 Then
        ShouldSkipTableProperty = True
        Exit Function
    End If

    If IsNull(varValue) Then
        ShouldSkipTableProperty = True
        Exit Function
    End If

    ' Skip internal metadata
    Select Case strName
        Case "NameMap", "Filter", "OrderBy"
            ShouldSkipTableProperty = True
            Exit Function
    End Select

    ' Skip table/query properties at their default values
    Select Case strName
        Case "Orientation":     ShouldSkipTableProperty = (varValue = 0)
        Case "DefaultView":     ShouldSkipTableProperty = (varValue = 2)
        Case "DisplayViewsOnSharePointSite": ShouldSkipTableProperty = (varValue = 1)
        Case "TotalsRow":       ShouldSkipTableProperty = IsDefaultBool(varValue, False)
        Case "FilterOnLoad":    ShouldSkipTableProperty = IsDefaultBool(varValue, False)
        Case "OrderByOn":       ShouldSkipTableProperty = IsDefaultBool(varValue, False)
        Case "OrderByOnLoad":   ShouldSkipTableProperty = IsDefaultBool(varValue, True)
        Case "HideNewField":    ShouldSkipTableProperty = IsDefaultBool(varValue, False)
        Case "ReadOnlyWhenDisconnected": ShouldSkipTableProperty = IsDefaultBool(varValue, False)
        Case "PublishToWeb":    ShouldSkipTableProperty = (varValue = 1)
        Case "Description"
            ShouldSkipTableProperty = (Len(Nz(varValue)) = 0)
        Case "BackTint", "BackShade", "AlternateBackTint", "AlternateBackShade", _
             "BorderTint", "BorderShade", "ForeTint", "ForeShade"
            ShouldSkipTableProperty = (IsNumeric(varValue) And Abs(CSng(varValue) - 100!) < 0.01)
        Case "ThemeFontIndex", "AlternateBackThemeColorIndex", _
             "DatasheetGridlinesThemeColorIndex", "DatasheetForeThemeColorIndex", _
             "BackThemeColorIndex", "BorderThemeColorIndex", "ForeThemeColorIndex"
            ShouldSkipTableProperty = (varValue = -1)
        Case "DatasheetBackColor"
            ShouldSkipTableProperty = (varValue = 16777215)
        Case "DatasheetBackColor12"
            ShouldSkipTableProperty = (varValue = -2147483643)
        Case "DatasheetAlternateBackColor"
            ShouldSkipTableProperty = (varValue = 15523798)
        ' Query-specific defaults
        Case "ODBCTimeout":     ShouldSkipTableProperty = (varValue = 60)
        Case "MaxRecords":      ShouldSkipTableProperty = (varValue = 0)
        Case "RecordsetType":   ShouldSkipTableProperty = (varValue = 0)
        Case "UseTransaction":  ShouldSkipTableProperty = IsDefaultBool(varValue, True)
        Case "RecordLocks":     ShouldSkipTableProperty = (varValue = 0)
        Case "RowHeight":       ShouldSkipTableProperty = (varValue = 65535 Or varValue = -1)
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsDefaultBool
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Helper to compare boolean values from the blob (which may be stored as
'           : numeric 0/1 or as actual Boolean) against expected defaults.
'---------------------------------------------------------------------------------------
'
Private Function IsDefaultBool(varValue As Variant, blnDefault As Boolean) As Boolean
    If IsNumeric(varValue) Then
        If blnDefault Then
            IsDefaultBool = (varValue <> 0)
        Else
            IsDefaultBool = (varValue = 0 Or varValue = False)
        End If
    ElseIf VarType(varValue) = vbBoolean Then
        IsDefaultBool = (CBool(varValue) = blnDefault)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadByte
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Read one byte and advance cursor.
'---------------------------------------------------------------------------------------
'
Private Function ReadByte() As Byte
    If m_lngPos < m_lngSize + LBound(m_bteData) Then
        ReadByte = m_bteData(m_lngPos)
        m_lngPos = m_lngPos + 1
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadInt
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Read a 2-byte little-endian unsigned integer and advance cursor.
'---------------------------------------------------------------------------------------
'
Private Function ReadInt() As Long
    If m_lngPos + 1 < m_lngSize + LBound(m_bteData) Then
        ReadInt = CLng(m_bteData(m_lngPos)) + CLng(m_bteData(m_lngPos + 1)) * 256
        m_lngPos = m_lngPos + 2
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadLong
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Read a 4-byte little-endian signed Long and advance cursor.
'           : Uses intermediate Long to avoid overflow with large unsigned values.
'---------------------------------------------------------------------------------------
'
Private Function ReadLong() As Long
    Dim lng0 As Long, lng1 As Long, lng2 As Long, lng3 As Long
    If m_lngPos + 3 < m_lngSize + LBound(m_bteData) Then
        lng0 = m_bteData(m_lngPos)
        lng1 = m_bteData(m_lngPos + 1)
        lng2 = m_bteData(m_lngPos + 2)
        lng3 = m_bteData(m_lngPos + 3)
        ' Reconstruct as signed Long, handling the sign bit manually
        If lng3 < 128 Then
            ReadLong = lng0 + lng1 * &H100& + lng2 * &H10000 + lng3 * &H1000000
        Else
            ReadLong = lng0 + lng1 * &H100& + lng2 * &H10000 + (lng3 - 256) * &H1000000
        End If
        m_lngPos = m_lngPos + 4
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadUtf16
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Read lngByteLen bytes of UTF-16LE text and return as VBA String.
'---------------------------------------------------------------------------------------
'
Private Function ReadUtf16(lngByteLen As Long) As String
    Dim lngCharCount As Long
    Dim lngChar As Long
    If lngByteLen <= 0 Or m_lngPos + lngByteLen > m_lngSize + LBound(m_bteData) Then Exit Function
    lngCharCount = lngByteLen \ 2
    ReadUtf16 = String$(lngCharCount, vbNullChar)
    For lngChar = 1 To lngCharCount
        Mid$(ReadUtf16, lngChar, 1) = ChrW$(CLng(m_bteData(m_lngPos)) + CLng(m_bteData(m_lngPos + 1)) * 256)
        m_lngPos = m_lngPos + 2
    Next lngChar
End Function
