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
    ReDim astrPages(0 To clngInitialPages - 1) As String
    
    ' Prepare first page
    astrPages(0) = Space$(clngPageSize)
    
End Sub


' Add to the string buffer
Public Sub Add(strAddString As String)

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
            If lngCurrentPos = clngPageSize Then
                ' See if we already have a new page available in the array
                If lngCurrentPage = UBound(astrPages) Then
                    ' Need to add a page to the array.
                    ReDim Preserve astrPages(0 To lngCurrentPage + 1)
                End If
                ' Prepare page as a buffer
                lngCurrentPage = lngCurrentPage + 1
                astrPages(lngCurrentPage) = Space$(clngPageSize)
                lngCurrentPos = 0
            End If
            
            ' See if it fits on the current page
            lngRemaining = clngPageSize - lngCurrentPos
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
                lngCurrentPos = clngPageSize
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
    lngTotalLen = lngCurrentPos + (lngCurrentPage * clngPageSize)
    
    ' We can't remove more characters than we put in the string to start with.
    If lngChars > lngTotalLen Then
        ' Go to beginning
        lngCurrentPage = 0
        lngCurrentPos = 1
    Else
        ' Get new absolute position
        lngNewPosition = lngTotalLen - lngChars
        ' Calculate full pages
        lngCurrentPage = (lngNewPosition \ clngPageSize)
        ' Set position on partial page
        lngCurrentPos = lngNewPosition - (lngCurrentPage * clngPageSize)
    End If
    
End Sub


' Returns the accumulated string
Public Function GetStr() As String

    Dim lngCnt As Long
    
    ' Prepare return string. This should be the filled pages plus the last
    ' partial page, divided by 2 to get the string length instead of byte length.
    GetStr = Space$((lngCurrentPage * clngPageSize) + lngCurrentPos)
    
    ' Loop through filled pages, overlaying on return string.
    ' (Last partial page is automatically trimmed based on returned string size.)
    If Len(GetStr) > 0 Then
        For lngCnt = 0 To lngCurrentPage
            Mid$(GetStr, (lngCnt * clngPageSize) + 1, clngPageSize) = astrPages(lngCnt)
        Next lngCnt
    End If
    
End Function


' returns the length of the string, based on the current position
' (Faster than building the string just to check the length)
Public Function Length() As Double
    Length = (lngCurrentPage * clngPageSize) + lngCurrentPos
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

    SetPageSize 10
    
    Add "abcdefghij"
    Add "k"
    Debug.Assert Len(GetStr) = 11
    Remove 2
    Debug.Assert Len(GetStr) = 9
    Add "jkl"
    Debug.Assert Len(GetStr) = 12
    Debug.Assert GetStr = "abcdefghijkl"
    Add "m123456789"
    Remove 11
    Debug.Assert GetStr = "abcdefghijk"
    
End Sub