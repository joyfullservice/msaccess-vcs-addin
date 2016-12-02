VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConcat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'String Concatenation Class
'
'Written By Nir Sofer
'http://nirsoft.mirrorz.com
'
'The following class can help you to Concatenate
'small strings in more efficient way than
'the standard VB "&" operator does it.
'
'In order to Concatenate strings with this class:
'
'1. Create a new instance of this class.
'2. Call the Add method to add small strings into the main string.
'3. Call the GetStr function in order to get the accumulated string.

Private strText           As String   'The stored string
Public lngAllocated       As Long     'Number of allocated characters
Public lngUsed            As Long     'Number of characters in use
Public lngAllocSize       As Long     'Number of characters to allocate every time

Private Sub Class_Initialize()
    'By default, allocate 1000 characters each time.
    'You can "play" with this value in order to get the best efficient string allocation for you.
    
    'If you increase this value, the string Concatenation operation
    'can be a little faster, but it'll waste more memory.
    
    'If you decrease this value, the string Concatenation operation
    'can be a little slower, but it'll waste less memory.
    lngAllocSize = 1000
End Sub

Public Sub Add(strAddString As String)
    Dim lngLen            As Long
    Dim lngToAllocate     As Long
    
    lngLen = Len(strAddString)
    If lngLen > 0 Then
        If lngUsed + lngLen > lngAllocated Then
            'Calculate the characters to allocate.
            lngToAllocate = lngAllocSize * (1 + (lngUsed + lngLen - lngAllocated) \ lngAllocSize)
            'Allocate more space in the string.
            strText = strText & String$(lngToAllocate, " ")
            lngAllocated = lngAllocated + lngToAllocate
        End If
        
        Mid$(strText, lngUsed + 1, lngLen) = strAddString
        lngUsed = lngUsed + lngLen
    End If
End Sub

'Returns the accumulated string
Public Function GetStr() As String
    GetStr = Mid$(strText, 1, lngUsed)
End Function
