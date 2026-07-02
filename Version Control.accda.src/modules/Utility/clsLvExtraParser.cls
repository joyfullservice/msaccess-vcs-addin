VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLvExtraParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsLvExtraParser
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Parses the undocumented LvExtra binary blob from MSysObjects to extract
'           : query design layout data (window position, designer pane dimensions,
'           : table positions within the designer).
'           :
'           : LvExtra is only present for queries last saved in Design View. It stores
'           : the visual layout of tables/queries in the query designer and various
'           : UI state properties.
'           :
'           : Binary format (reverse-engineered from live Access databases):
'           :   [4: magic 0x99 0x99 0xCE 0xAC] [12: padding 0xAA]
'           :   [16: window RECT (L,T,R,B as Longs)]
'           :   [4: state Long]
'           :   [16: designer pane RECT (L,T,R,B as Longs)]
'           :   [8: grid origin (L,T as Longs)]
'           :   [4: ColumnsShown Long]
'           :   [4: table count Long]
'           :   [per table entry: 284 bytes fixed]
'           :     20 bytes: 5 Longs (L,T,R,B, scrollTop)
'           :     132 bytes: name (null-terminated UTF-16LE, zero-padded to 132 bytes)
'           :     132 bytes: alias (null-terminated UTF-16LE, zero-padded to 132 bytes)
'           :   There is always one extra phantom entry (all zeros) beyond tableCount.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Utility")

Private m_bteData() As Byte
Private m_lngPos As Long
Private m_lngSize As Long

Private Const MAGIC_B0 As Byte = &H99
Private Const MAGIC_B1 As Byte = &H99
Private Const MAGIC_B2 As Byte = &HCE
Private Const MAGIC_B3 As Byte = &HAC
Private Const HEADER_SIZE As Long = 16
Private Const TABLE_NAME_FIELD_SIZE As Long = 132
Private Const TABLE_ENTRY_SIZE As Long = 284


'---------------------------------------------------------------------------------------
' Procedure : ParseLvExtra
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Main entry point. Parses an LvExtra binary blob and returns a Dictionary
'           : containing the design layout data suitable for JSON serialization and
'           : .qdef generation.
'---------------------------------------------------------------------------------------
'
Public Function ParseLvExtra(bteData() As Byte) As Dictionary

    Dim dResult As New Dictionary
    Dim dWindow As Dictionary
    Dim dPane As Dictionary
    Dim dTable As Dictionary
    Dim colTables As Collection
    Dim lngTableCount As Long
    Dim lngIdx As Long

    m_bteData = bteData
    m_lngSize = UBound(m_bteData) - LBound(m_bteData) + 1
    m_lngPos = LBound(m_bteData)

    ' Validate magic header
    If m_lngSize < 68 Then GoTo Done
    If m_bteData(m_lngPos) <> MAGIC_B0 Or m_bteData(m_lngPos + 1) <> MAGIC_B1 Or _
       m_bteData(m_lngPos + 2) <> MAGIC_B2 Or m_bteData(m_lngPos + 3) <> MAGIC_B3 Then GoTo Done

    ' Skip header (magic + padding)
    m_lngPos = LBound(m_bteData) + HEADER_SIZE

    ' Window RECT
    Set dWindow = New Dictionary
    dWindow.Add "left", ReadLong()
    dWindow.Add "top", ReadLong()
    dWindow.Add "right", ReadLong()
    dWindow.Add "bottom", ReadLong()
    dResult.Add "window", dWindow

    ' State
    dResult.Add "state", ReadLong()

    ' Designer pane RECT
    Set dPane = New Dictionary
    dPane.Add "left", ReadLong()
    dPane.Add "top", ReadLong()
    dPane.Add "right", ReadLong()
    dPane.Add "bottom", ReadLong()
    dResult.Add "designerPane", dPane

    ' Grid origin
    dResult.Add "gridLeft", ReadLong()
    dResult.Add "gridTop", ReadLong()

    ' Columns shown
    dResult.Add "columnsShown", ReadLong()

    ' Table count
    lngTableCount = ReadLong()

    ' Table entries: fixed 284-byte records (5 Longs + 132-byte name + 132-byte alias).
    Set colTables = New Collection
    Dim lngBlobEnd As Long
    lngBlobEnd = m_lngSize + LBound(m_bteData)
    For lngIdx = 1 To lngTableCount
        If m_lngPos + TABLE_ENTRY_SIZE > lngBlobEnd Then Exit For
        Set dTable = New Dictionary
        dTable.Add "left", ReadLong()
        dTable.Add "top", ReadLong()
        dTable.Add "right", ReadLong()
        dTable.Add "bottom", ReadLong()
        dTable.Add "scrollTop", ReadLong()
        dTable.Add "name", ReadFixedUtf16(TABLE_NAME_FIELD_SIZE)
        dTable.Add "alias", ReadFixedUtf16(TABLE_NAME_FIELD_SIZE)
        If Len(dTable("name")) > 0 Then
            colTables.Add dTable
        End If
    Next lngIdx
    dResult.Add "tables", colTables

Done:
    Set ParseLvExtra = dResult

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadLong
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Read a 4-byte little-endian signed Long and advance cursor.
'---------------------------------------------------------------------------------------
'
Private Function ReadLong() As Long
    Dim lng0 As Long, lng1 As Long, lng2 As Long, lng3 As Long
    If m_lngPos + 3 < m_lngSize + LBound(m_bteData) Then
        lng0 = m_bteData(m_lngPos)
        lng1 = m_bteData(m_lngPos + 1)
        lng2 = m_bteData(m_lngPos + 2)
        lng3 = m_bteData(m_lngPos + 3)
        If lng3 < 128 Then
            ReadLong = lng0 + lng1 * &H100& + lng2 * &H10000 + lng3 * &H1000000
        Else
            ReadLong = lng0 + lng1 * &H100& + lng2 * &H10000 + (lng3 - 256) * &H1000000
        End If
        m_lngPos = m_lngPos + 4
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadNullTermUtf16
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Read a null-terminated UTF-16LE string from the blob. Stops at the first
'           : null character (0x00 0x00) and advances past it.
'---------------------------------------------------------------------------------------
'
Private Function ReadNullTermUtf16() As String

    Dim lngStart As Long
    Dim lngCharCount As Long
    Dim lngEnd As Long
    Dim lngChar As Long

    lngStart = m_lngPos
    lngEnd = m_lngSize + LBound(m_bteData)

    ' Scan for null terminator (two zero bytes on even boundary)
    lngCharCount = 0
    Do While m_lngPos + 1 < lngEnd
        If m_bteData(m_lngPos) = 0 And m_bteData(m_lngPos + 1) = 0 Then
            ' Found null terminator
            m_lngPos = m_lngPos + 2
            Exit Do
        End If
        lngCharCount = lngCharCount + 1
        m_lngPos = m_lngPos + 2
    Loop

    ' Build string from the chars we found
    If lngCharCount > 0 Then
        ReadNullTermUtf16 = String$(lngCharCount, vbNullChar)
        For lngChar = 1 To lngCharCount
            Mid$(ReadNullTermUtf16, lngChar, 1) = ChrW$(CLng(m_bteData(lngStart + (lngChar - 1) * 2)) + _
                CLng(m_bteData(lngStart + (lngChar - 1) * 2 + 1)) * 256)
        Next lngChar
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReadFixedUtf16
' Author    : Adam Waller
' Date      : 4/10/2026
' Purpose   : Read a null-terminated UTF-16LE string from a fixed-size field. Reads
'           : characters until a null terminator, then advances the cursor by the full
'           : field size regardless of string length.
'---------------------------------------------------------------------------------------
'
Private Function ReadFixedUtf16(lngFieldSize As Long) As String

    Dim lngStart As Long
    Dim lngCharCount As Long
    Dim lngMaxChars As Long
    Dim lngChar As Long
    Dim lngI As Long

    lngStart = m_lngPos
    lngMaxChars = lngFieldSize \ 2

    ' Scan for null terminator within the fixed field
    lngCharCount = 0
    For lngI = 0 To lngMaxChars - 1
        If m_bteData(lngStart + lngI * 2) = 0 And m_bteData(lngStart + lngI * 2 + 1) = 0 Then
            Exit For
        End If
        lngCharCount = lngCharCount + 1
    Next lngI

    ' Build string from the chars we found
    If lngCharCount > 0 Then
        ReadFixedUtf16 = String$(lngCharCount, vbNullChar)
        For lngChar = 1 To lngCharCount
            Mid$(ReadFixedUtf16, lngChar, 1) = ChrW$(CLng(m_bteData(lngStart + (lngChar - 1) * 2)) + _
                CLng(m_bteData(lngStart + (lngChar - 1) * 2 + 1)) * 256)
        Next lngChar
    End If

    ' Always advance by the full field size
    m_lngPos = m_lngPos + lngFieldSize

End Function
