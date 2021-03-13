Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' String concatenation (joining strings together) can have a significant
' performance impact when you are using the ampersand character to join
' strings together. While negligible in occasional use, if you start
' running tens of thousands of these in a loop, it can really bog
' down the processing due to the memory reallocations happening behind
' the scenes. In those cases it is better to use the Mid$() function to
' change an existing buffer to build the return string.

' Special thanks to Nir Sofer - http://www.nirsoft.net/vb/strclass.html
' and Chris Lucas - http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=37141&lngWId=1
' for their inspiration with these concepts.

' Set this to any character or string to add after each
' call to `.Add()`. A common example would be vbCrLf.
Public AppendOnAdd As String

' Set up an array of pages to hold strings
Private astrPages() As String
Private lngCurrentPage As Long
Private lngCurrentPos As Long
Private lngPageSize As Long
Private lngInitialPages As Long

' These defaults can be tweaked as needed
Const clngPageSize As Long = 4096
Const clngInitialPages As Long = 100



' Prepares the initial buffer page
Private Sub Class_Initialize()
    
    If lngPageSize = 0 Then lngPageSize = clngPageSize
    If lngInitialPages = 0 Then lngInitialPages = clngInitialPages
    
    ' Set up the initial array of pages.
    ReDim astrPages(0 To lngInitialPages - 1) As String
    
    ' Prepare first page
    astrPages(0) = Space$(lngPageSize)
    
End Sub


' Add 1 or more strings (avoiding the string conversion of paramarray)
Public Sub Add(str1 As String, Optional str2 As String, Optional str3 As String, Optional str4 As String, Optional str5 As String, _
    Optional str6 As String, Optional str7 As String, Optional str8 As String, Optional str9 As String, Optional str10 As String)
    If str1 <> vbNullString Then AddString str1
    If str2 <> vbNullString Then AddString str2
    If str3 <> vbNullString Then AddString str3
    If str4 <> vbNullString Then AddString str4
    If str5 <> vbNullString Then AddString str5
    If str6 <> vbNullString Then AddString str6
    If str7 <> vbNullString Then AddString str7
    If str8 <> vbNullString Then AddString str8
    If str9 <> vbNullString Then AddString str9
    If str10 <> vbNullString Then AddString str10
    AddString AppendOnAdd
End Sub


' Add to the string buffer
Private Sub AddString(strAddString As String)

    Dim lngLen          As Long
    Dim lngRemaining    As Long
    Dim lngAddStrPos    As Long
    Dim lngAddLen       As Long
    
    ' Get length of new string
    lngLen = Len(strAddString)
    
    ' No need to process a zero-length string
    If lngLen > 0 Then
        ' Set starting position for string we are adding
        lngAddStrPos = 1
        
        ' Continue filling pages till we reach the end of the new string
        Do While lngAddStrPos <= lngLen
            
            ' Check to see if we need a new page
            If lngCurrentPos = lngPageSize Then
                ' See if we already have a new page available in the array
                If lngCurrentPage = UBound(astrPages) Then
                    ' Need to add a page to the array.
                    ReDim Preserve astrPages(0 To lngCurrentPage + 1)
                End If
                ' Prepare page as a buffer
                lngCurrentPage = lngCurrentPage + 1
                astrPages(lngCurrentPage) = Space$(lngPageSize)
                lngCurrentPos = 0
            End If
            
            ' See if it fits on the current page
            lngRemaining = lngPageSize - lngCurrentPos
            If (lngLen - (lngAddStrPos - 1)) <= lngRemaining Then
                ' Yes, add to current page.
                lngAddLen = (lngLen - (lngAddStrPos - 1))
                Mid$(astrPages(lngCurrentPage), lngCurrentPos + 1, lngAddLen) = Mid$(strAddString, lngAddStrPos)
                lngAddStrPos = lngLen + 1
                lngCurrentPos = lngCurrentPos + lngAddLen
            Else
                ' Fill remaining available space on current page.
                Mid$(astrPages(lngCurrentPage), lngCurrentPos + 1, lngRemaining) = Mid$(strAddString, lngAddStrPos, lngRemaining)
                ' Note position in new string
                lngCurrentPos = lngPageSize
                lngAddStrPos = lngAddStrPos + lngRemaining
            End If
        
        ' Move to next page, if needed
        Loop
    
    End If
    
End Sub


' Removes the specified number of chacters from the string.
' (Technically just moves the position back)
Public Sub Remove(lngChars As Long)
    
    Dim lngTotalLen As Long
    Dim lngNewPosition As Long
    
    ' Get total length of current string including all pages
    lngTotalLen = lngCurrentPos + (lngCurrentPage * lngPageSize)
    
    ' We can't remove more characters than we put in the string to start with.
    If lngChars > lngTotalLen Then
        ' Go to beginning
        lngCurrentPage = 0
        lngCurrentPos = 1
    Else
        ' Get new absolute position
        lngNewPosition = lngTotalLen - lngChars
        ' Calculate full pages
        lngCurrentPage = (lngNewPosition \ lngPageSize)
        ' Set position on partial page
        lngCurrentPos = lngNewPosition - (lngCurrentPage * lngPageSize)
    End If
    
End Sub


' Returns the accumulated string
Public Function GetStr() As String

    Dim lngCnt As Long
    
    ' Prepare return string. This should be the filled pages plus the last
    ' partial page, divided by 2 to get the string length instead of byte length.
    GetStr = Space$((lngCurrentPage * lngPageSize) + lngCurrentPos)
    
    ' Loop through filled pages, overlaying on return string.
    ' (Last partial page is automatically trimmed based on returned string size.)
    If Len(GetStr) > 0 Then
        For lngCnt = 0 To lngCurrentPage
            Mid$(GetStr, (lngCnt * lngPageSize) + 1, lngPageSize) = astrPages(lngCnt)
        Next lngCnt
    End If
    
End Function


' Return a partial string from a specified position
Public Function MidStr(lngStart As Long, Optional lngLength)

    Dim lngPage As Long
    Dim lngPos As Long
    Dim lngStartPage As Long
    Dim lngStartPos As Long
    
    ' Prepare return string length.
    If IsMissing(lngLength) Then
        ' Return remaining string after lngStart
        lngLength = (Me.Length - lngStart) + 1
        MidStr = Space$(lngLength)
    Else
        ' Return a specified number of characters
        MidStr = Space$(lngLength)
    End If
    
    ' Determine start page and position for return string
    lngStartPage = (lngStart - 1) \ lngPageSize ' Zero based page
    lngStartPos = lngStart - (lngStartPage * lngPageSize)
    
    ' Loop through filled pages, overlaying on return string.
    ' (Last partial page is automatically trimmed based on returned string size.)
    If Len(MidStr) > 0 Then
        For lngPage = lngStartPage To lngCurrentPage
            ' Could start at any point on first page
            If lngPage = lngStartPage Then
                Mid$(MidStr, 1) = Mid$(astrPages(lngPage), lngStartPos)
                ' lngPos is the current position in the new string
                lngPos = lngPageSize - (lngStartPos - 2)
            Else
                ' Pull whole pages as needed
                Mid$(MidStr, lngPos) = astrPages(lngPage)
                lngPos = lngPos + lngPageSize
            End If
            ' Exit when we have filled the requested string.
            If lngPos > lngLength Then Exit For
        Next lngPage
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : Right
' Author    : Adam Waller
' Date      : 11/5/2020
' Purpose   : Return the rightmost specified number of characters.
'---------------------------------------------------------------------------------------
'
Public Function RightStr(lngLength As Long) As String
    If Me.Length > lngLength Then
        RightStr = MidStr((Me.Length - lngLength) + 1)
    Else
        RightStr = GetStr
    End If
End Function


' returns the length of the string, based on the current position
' (Faster than building the string just to check the length)
Public Function Length() As Double
    Length = (lngCurrentPage * lngPageSize) + lngCurrentPos
End Function


' Reset the buffer without changing the page size
Public Sub Clear()
    
    Class_Initialize
    
    ' Reset positions
    lngCurrentPage = 0
    lngCurrentPos = 0

End Sub


' Manually set page size if you want something different from the default.
Public Sub SetPageSize(lngNewPageSize As Long, Optional lngNewInitialPages As Long)
    If lngCurrentPage > 0 Or lngCurrentPos > 1 Then
        MsgBox "Please set the page size before adding any data", vbExclamation, "Error in clsConcat"
    Else
        lngPageSize = lngNewPageSize
        If lngNewInitialPages > 0 Then lngInitialPages = lngNewInitialPages
        ' Reinitialize with the updated sizes
        Class_Initialize
    End If
End Sub


' Test the class to make sure we are paging correctly.
Public Sub SelfTest()

    SetPageSize 10, 5
    
    Debug.Assert UBound(astrPages) = 4
    Add "abcdefghij"
    Add "k"
    Debug.Assert Len(GetStr) = 11
    Debug.Assert Length = 11
    Remove 2
    Debug.Assert Len(GetStr) = 9
    Add "jkl"
    Debug.Assert Len(GetStr) = 12
    Debug.Assert GetStr = "abcdefghijkl"
    Add "m123456789"
    Remove 11
    Debug.Assert GetStr = "abcdefghijk"
    Debug.Assert MidStr(1, 1) = "a"
    Debug.Assert MidStr(11, 1) = "k"
    Debug.Assert MidStr(2, 3) = "bcd"
    Debug.Assert MidStr(8) = "hijk"
    Debug.Assert MidStr(10, 1) = "j"
    Debug.Assert RightStr(1) = "k"
    Debug.Assert RightStr(100) = "abcdefghijk"
    
    ' Verify paging
    With Me
        .Clear
        .SetPageSize 5, 2
        .Add "1234"
        Debug.Assert .GetStr = "1234"
        .Add "5"
        Debug.Assert .GetStr = "12345"
        .Add "6"
        Debug.Assert .GetStr = "123456"
        .Add "789"
        Debug.Assert .GetStr = "123456789"
        .Add "0"
        Debug.Assert .GetStr = "1234567890"
        .Add "A"
        Debug.Assert .GetStr = "1234567890A"
        .Remove 1
        Debug.Assert .GetStr = "1234567890"
    End With
    
End Sub