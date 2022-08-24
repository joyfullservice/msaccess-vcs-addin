Attribute VB_Name = "modStringFunctions"
'---------------------------------------------------------------------------------------
' Module    : modStringFunctions
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : General functions that don't fit more specifically into another module.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

Private Const ModuleName As String = "modStringFunctions"


'---------------------------------------------------------------------------------------
' Procedure : MultiReplace
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Does a string replacement of multiple items in one call.
'---------------------------------------------------------------------------------------
'
Public Function MultiReplace(ByVal strText As String _
                            , ParamArray varPairs()) As String
    Dim intPair As Long
    For intPair = 0 To UBound(varPairs) Step 2
        strText = Replace(strText, varPairs(intPair), varPairs(intPair + 1))
    Next intPair
    MultiReplace = strText
End Function


'---------------------------------------------------------------------------------------
' Procedure : Repeat
' Author    : Adam Waller
' Date      : 4/29/2021
' Purpose   : Repeat a string a specified number of times
'---------------------------------------------------------------------------------------
'
Public Function Repeat(strText As String _
                    , lngTimes As Long) As String
    Repeat = Replace$(Space$(lngTimes), " ", strText)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Coalesce
' Author    : Adam Waller
' Date      : 5/15/2021
' Purpose   : Return the first non-empty string from an array of string values
'---------------------------------------------------------------------------------------
'
Public Function Coalesce(ParamArray varStrings()) As String
    Dim intString As Long
    For intString = 0 To UBound(varStrings)
        If Nz(varStrings(intString)) <> vbNullString Then
            Coalesce = varStrings(intString)
            Exit For
        End If
    Next intString
End Function


'---------------------------------------------------------------------------------------
' Procedure : PadRight
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Pads a string
'---------------------------------------------------------------------------------------
'
Private Function PadRight(strText As String _
                        , lngLen As Long _
                        , Optional lngMinTrailingSpaces As Long = 1) As String

    Dim strResult As String
    Dim strTrimmed As String
    
    strResult = Space$(lngLen)
    strTrimmed = Left$(strText, lngLen - lngMinTrailingSpaces)
    
    ' Use mid function to write over existing string of spaces.
    Mid$(strResult, 1, Len(strTrimmed)) = strTrimmed
    PadRight = strResult
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : ListResult
' Author    : Adam Waller, hecon5
' Date      : 11/3/2020, May 18, 2022
' Purpose   : List the result of a test in a fixed width format. The result strings
'           : are positioned at the number of characters specified.
'           : I.e:
'           : MyFancyTest      23     2.45
'---------------------------------------------------------------------------------------
'
Public Function ListResult(ByRef strHeading As String _
                            , ByRef strResult1 As String _
                            , ByRef strResult2 As String _
                            , ByRef lngCol() As Long) As String
    ListResult = ListResultIndent(strHeading, strResult1, strResult2, lngCol)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ListResultIndent
' Author    : hecon5
' Date      : May 18, 2022
' Purpose   : List the result of a test in a fixed width format. The result strings
'           : are positioned at the number of characters specified. strHeadings longer
'           : than the column are repeated on the following line.
'           : I.e:
'           :   MyFancyTest                   23        2.45
'           :   Short Category                1         6.83
'           :   Long Category name from long  1         1.14
'           :       town
'           :   Some Other Category           2         1000
'---------------------------------------------------------------------------------------
'
Public Function ListResultIndent(ByRef strHeading As String _
                                , ByRef strResult1 As String _
                                , ByRef strResult2 As String _
                                , ByRef lngCol() As Long _
                                , Optional ByVal ColumnIndent As Long = 4) As String
    
    Dim Col1StrArr() As String
    Dim Col2StrArr() As String
    Dim Col3StrArr() As String
    
    Dim Col1Rows As Long
    Dim Col2Rows As Long
    Dim Col3Rows As Long
    
    Dim RowTotal As Long
    Dim RowPosition As Long
    
    Dim StrOutput As clsConcat
    
    On Error Resume Next
    Perf.OperationStart ModuleName & ".ListResultIndent"
    
    Col1StrArr = FitStringToColumn(strHeading, lngCol(0) - 1, ColumnIndent)
    Col2StrArr = FitStringToColumn(strResult1, lngCol(1) - 1, ColumnIndent)
    Col3StrArr = FitStringToColumn(strResult2, lngCol(2) - 1, ColumnIndent)
    
    Col1Rows = UBound(Col1StrArr)
    Col2Rows = UBound(Col2StrArr)
    Col3Rows = UBound(Col3StrArr)
    
    RowTotal = MaxValue(Col1Rows, Col2Rows, Col3Rows)
        
    Set StrOutput = New clsConcat
    
    For RowPosition = 0 To RowTotal
    
        If Col1Rows >= RowPosition Then
            StrOutput.Add PadRight(Col1StrArr(RowPosition), lngCol(0))
        Else
            StrOutput.Add Space$(lngCol(0))
        End If
        If Col2Rows >= RowPosition Then
            StrOutput.Add PadRight(Col2StrArr(RowPosition), lngCol(1))
        Else
            StrOutput.Add Space$(lngCol(1))
        End If
        If Col3Rows >= RowPosition Then
            StrOutput.Add PadRight(Col3StrArr(RowPosition), lngCol(2))
        Else
            StrOutput.Add Space$(lngCol(2))
        End If
        ' Don't add a new line for the last line; it's handled outside this tool
        If RowTotal > RowPosition Then StrOutput.Add vbNewLine
    
    Next RowPosition

    ListResultIndent = StrOutput.GetStr
    Perf.OperationEnd
End Function


'---------------------------------------------------------------------------------------
' Procedure : FitStringToColumn
' Author    : hecon5
' Date      : May 18, 2022
' Purpose   : Takes in a long string and returns an array of strings ColumnWidth wide.
'---------------------------------------------------------------------------------------
'
Public Function FitStringToColumn(ByRef LongString As String _
                                , Optional ByRef ColumnWidth As Long = 200 _
                                , Optional ByRef ColumnIndent As Long = 0) As String()

    Dim RowTotal As Long
    Dim StrLen As Long
    Dim StrIndentedLen As Long
    Dim StrTextWidth As Long
    Dim StrPosition As Long
    Dim ArrPosition As Long
    Dim StrArr() As String
    Dim ColumnWidthInternal As Long
    
    On Error Resume Next
    Perf.OperationStart ModuleName & ".FitStringToColumn"
    If Len(LongString) = 0 Then Exit Function
    ColumnWidthInternal = ColumnWidth
    If ColumnWidthInternal <= 0 Then ColumnWidthInternal = 1
    
    StrTextWidth = ColumnWidthInternal - ColumnIndent
    
    StrLen = Len(LongString)
    RowTotal = RoundUp((StrLen - ColumnWidthInternal) / StrTextWidth) + 1
    If RowTotal < 1 Then RowTotal = 1
    StrPosition = 1
    
    ReDim StrArr(0 To (RowTotal - 1))
    
    ' The first row is longer.
    StrArr(ArrPosition) = mid$(LongString, StrPosition, ColumnWidthInternal)
    If RowTotal <= 1 Then GoTo Exit_Here ' Don't do the rest if there's only one row...
    
    StrPosition = StrPosition + ColumnWidthInternal

    For ArrPosition = 1 To (RowTotal - 1)
        StrArr(ArrPosition) = Space$(ColumnIndent) & mid$(LongString, StrPosition, StrTextWidth)
        StrPosition = StrPosition + StrTextWidth
    Next ArrPosition

Exit_Here:
    CatchAny eelError, "Could not fit to column", Perf.CurrentOperationName
    FitStringToColumn = StrArr
    Perf.OperationEnd
End Function


'---------------------------------------------------------------------------------------
' Procedure : FitStringToWidth
' Author    : hecon5
' Date      : May 18, 2022
' Purpose   : Takes in a long string and returns a string DesiredWidth wide if longer than MaxWidth.
'             Useful when used on a text box where you'd prefer a specific width, but
'             allowing a slightly longer string would potentially be more appealing to users.
'             This avoids a 1-2 word dangling line.
'---------------------------------------------------------------------------------------
'
Public Function FitStringToWidth(ByRef LongString As String _
                                , Optional ByRef MaxWidth As Long = 200 _
                                , Optional ByRef DesiredWidth As Long = 75) As String
    ' Fits a string to a message box if it's wider than MaxWidth
    Dim OutputConcat As clsConcat
    Dim StrPosition As Long
    Dim StrLen As Long ' Length of total string
    Dim NewLineCount As Long ' Number of newlines
    Dim ArrPosition As Long
    Dim StrArrLen As Long ' Length of substring
    Dim StringArr() As String
    
    Perf.OperationStart "FitStringToWidth"
    StrLen = Len(LongString)
    If StrLen > MaxWidth Then
        Perf.OperationStart "FitStringToWidth.Resize"
        StringArr = Split(LongString, vbNewLine, , vbTextCompare)
        NewLineCount = UBound(StringArr) - LBound(StringArr)
        Set OutputConcat = New clsConcat
        For ArrPosition = 0 To NewLineCount
            StrPosition = 1
            StrArrLen = Len(StringArr(ArrPosition))
            If ArrPosition > 0 Then OutputConcat.Add vbNewLine
            Do While StrPosition < StrArrLen
                If StrPosition > 1 Then OutputConcat.Add vbNewLine
                OutputConcat.Add mid$(StringArr(ArrPosition), StrPosition, DesiredWidth)
                StrPosition = StrPosition + DesiredWidth
            Loop
        Next ArrPosition
        FitStringToWidth = OutputConcat.GetStr
        Perf.OperationEnd
    Else
        FitStringToWidth = LongString
    End If
    Perf.OperationEnd
End Function
