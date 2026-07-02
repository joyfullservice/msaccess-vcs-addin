VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormLayoutParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsFormLayoutParser
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Parses Access SaveAsText .form/.report files into a tree of clsLayoutNode
'           : objects for SVG layout generation. Handles Begin/End block nesting,
'           : quoted and unquoted property values, multi-line hex data blocks
'           : (ImageData, RecSrcDt), and stops at the CodeBehindForm boundary.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Core")

Private Type udtThis
    strContent As String
    aLines() As String
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : Parse
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Parse the form/report source file content and return the root node.
'---------------------------------------------------------------------------------------
'
Public Function Parse(strContent As String) As clsLayoutNode

    Dim cRoot As clsLayoutNode
    Dim colStack As Collection
    Dim cCurrent As clsLayoutNode
    Dim lngLine As Long
    Dim strLine As String
    Dim strTrimmed As String
    Dim strKey As String
    Dim strValue As String
    Dim lngPos As Long
    Dim blnInHexBlock As Boolean
    Dim strHexKey As String
    Dim cHexConcat As clsConcat

    this.strContent = strContent
    this.aLines = Split(strContent, vbCrLf)

    Set cRoot = New clsLayoutNode
    cRoot.ControlType = "Root"

    Set colStack = New Collection
    colStack.Add cRoot
    Set cCurrent = cRoot

    For lngLine = 0 To UBound(this.aLines)
        strLine = this.aLines(lngLine)
        strTrimmed = Trim$(strLine)

        ' Stop parsing at the VBA code boundary
        If strTrimmed = "CodeBehindForm" Then Exit For

        ' Handle multi-line hex data blocks (ImageData, RecSrcDt)
        If blnInHexBlock Then
            If strTrimmed = "End" Then
                ' Close hex block; store accumulated hex data
                If Not cHexConcat Is Nothing Then
                    cCurrent.Props(strHexKey) = cHexConcat.GetStr
                    Set cHexConcat = Nothing
                End If
                blnInHexBlock = False
            Else
                ' Accumulate hex line (strip trailing comma/space)
                If Not cHexConcat Is Nothing Then
                    cHexConcat.Add Replace(Replace(strTrimmed, ",", ""), " ", "")
                End If
            End If
            GoTo NextLine
        End If

        ' Skip blank lines
        If Len(strTrimmed) = 0 Then GoTo NextLine

        ' Check for Begin block
        If Left$(strTrimmed, 6) = "Begin " Then
            Dim cChild As clsLayoutNode
            Set cChild = New clsLayoutNode
            cChild.ControlType = Mid$(strTrimmed, 7)
            cCurrent.AddChild cChild
            colStack.Add cChild
            Set cCurrent = cChild
            GoTo NextLine
        End If

        ' Check for bare "Begin" (anonymous defaults block)
        If strTrimmed = "Begin" Then
            Dim cDefaults As clsLayoutNode
            Set cDefaults = New clsLayoutNode
            cDefaults.ControlType = "Defaults"
            cCurrent.AddChild cDefaults
            colStack.Add cDefaults
            Set cCurrent = cDefaults
            GoTo NextLine
        End If

        ' Check for End block
        If strTrimmed = "End" Then
            If colStack.Count > 1 Then
                colStack.Remove colStack.Count
                Set cCurrent = colStack(colStack.Count)
            End If
            GoTo NextLine
        End If

        ' Parse property: Key =Value (or Key = Value with spaces)
        lngPos = InStr(1, strTrimmed, " =")
        If lngPos > 0 Then
            strKey = Trim$(Left$(strTrimmed, lngPos - 1))
            strValue = Trim$(Mid$(strTrimmed, lngPos + 2))

            If strValue = "Begin" Then
                ' Start of a multi-line hex data block
                blnInHexBlock = True
                strHexKey = strKey
                Set cHexConcat = New clsConcat
                GoTo NextLine
            End If

            ' Strip surrounding quotes from string values
            If Len(strValue) >= 2 Then
                If Left$(strValue, 1) = """" And Right$(strValue, 1) = """" Then
                    strValue = Mid$(strValue, 2, Len(strValue) - 2)
                End If
            End If

            cCurrent.Props(strKey) = strValue
            GoTo NextLine
        End If

        ' Handle standalone flags like "PopUp = NotDefault" or just "PopUp"
        ' (properties with no " =" separator are flags set to the property name)
        If InStr(1, strTrimmed, "=") > 0 Then
            lngPos = InStr(1, strTrimmed, "=")
            strKey = Trim$(Left$(strTrimmed, lngPos - 1))
            strValue = Trim$(Mid$(strTrimmed, lngPos + 1))
            ' Strip quotes
            If Len(strValue) >= 2 Then
                If Left$(strValue, 1) = """" And Right$(strValue, 1) = """" Then
                    strValue = Mid$(strValue, 2, Len(strValue) - 2)
                End If
            End If
            cCurrent.Props(strKey) = strValue
        End If

NextLine:
    Next lngLine

    Set Parse = cRoot

End Function
