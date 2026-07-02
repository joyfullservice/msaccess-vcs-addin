VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormLayoutThemeColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsFormLayoutThemeColors
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Resolves Access theme color indices to RGB hex values by reading the
'           : exported .thmx theme file (OOXML ZIP containing theme/theme1.xml).
'           : Implements HSL-space tint/shade transforms per Microsoft spec.
'           : ThemeColorIndex mapping: 0=dk1, 1=lt1, 2=dk2, 3=lt2, 4-9=accent1-6,
'           : 10=hlink, 11=folHlink. Index -1 means "no theme" (use literal color).
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Core")

Private Type udtThis
    aBaseColors(0 To 11) As String  ' RGB hex strings (e.g. "4472C4")
    blnLoaded As Boolean
End Type
Private this As udtThis

' Slot names in OOXML order matching Access ThemeColorIndex values
Private Const THEME_SLOTS As String = "dk1,lt1,dk2,lt2,accent1,accent2,accent3,accent4,accent5,accent6,hlink,folHlink"


'---------------------------------------------------------------------------------------
' Procedure : LoadFromThemeFolder
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Finds the first .thmx file in the themes/ folder under the export root,
'           : extracts theme/theme1.xml, and parses the color scheme.
'---------------------------------------------------------------------------------------
'
Public Sub LoadFromThemeFolder(strExportFolder As String)

    Dim strThemesFolder As String
    Dim strThmxFile As String
    Dim strXmlPath As String
    Dim oFolder As Scripting.Folder
    Dim oSubFolder As Scripting.Folder
    Dim oFile As Scripting.File

    strThemesFolder = strExportFolder & "themes" & PathSep
    If Not FSO.FolderExists(strThemesFolder) Then Exit Sub
    Set oFolder = FSO.GetFolder(strThemesFolder)

    ' First, check for an already-extracted theme1.xml in any subfolder.
    ' This handles both manually extracted themes and our own cached extraction.
    For Each oSubFolder In oFolder.SubFolders
        strXmlPath = oSubFolder.Path & PathSep & "theme" & PathSep & "theme1.xml"
        If FSO.FileExists(strXmlPath) Then
            ParseThemeXml strXmlPath
            Exit Sub
        End If
    Next oSubFolder

    ' No extracted XML found. Look for a .thmx file and extract it.
    strThmxFile = vbNullString
    For Each oFile In oFolder.Files
        If LCase$(FSO.GetExtensionName(oFile.Name)) = "thmx" Then
            strThmxFile = oFile.Path
            Exit For
        End If
    Next oFile
    If Len(strThmxFile) = 0 Then Exit Sub

    ' Use a stable subfolder name derived from the .thmx filename so the
    ' extracted XML persists across exports and avoids repeated extraction.
    Dim strBaseName As String
    Dim strCacheFolder As String
    Dim strZipCopy As String
    strBaseName = FSO.GetBaseName(strThmxFile)
    strCacheFolder = GetTempFolder & "vcs_theme_" & strBaseName & PathSep
    strXmlPath = strCacheFolder & "theme" & PathSep & "theme1.xml"

    ' Check if we already extracted in a previous session
    If FSO.FileExists(strXmlPath) Then
        ParseThemeXml strXmlPath
        Exit Sub
    End If

    ' Extract theme1.xml from the .thmx (which is a ZIP archive).
    ' We copy just the "theme" folder rather than using ExtractFromZip,
    ' which has a polling timeout issue when the dest folder is non-empty.
    strZipCopy = strCacheFolder & "theme.zip"
    VerifyPath strZipCopy
    FSO.CopyFile strThmxFile, strZipCopy

    Dim oShell As Object
    Dim oZipNs As Object
    Dim oItem As Object
    Set oShell = CreateObject("Shell.Application")
    Set oZipNs = oShell.Namespace(CStr(strZipCopy))
    If Not oZipNs Is Nothing Then
        ' Find and extract the "theme" folder from the ZIP
        For Each oItem In oZipNs.Items
            If LCase$(oItem.Name) = "theme" Then
                oShell.Namespace(CStr(strCacheFolder)).CopyHere oItem, 16
                Exit For
            End If
        Next oItem
        ' Wait for async copy to complete (check for target file)
        Dim sngTimeout As Single
        sngTimeout = Timer + 10
        Do While Timer < sngTimeout
            If FSO.FileExists(strXmlPath) Then Exit Do
            Pause 0.1
        Loop
    End If

    If FSO.FileExists(strXmlPath) Then
        ParseThemeXml strXmlPath
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ParseThemeXml
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Parse theme1.xml to extract the 12 color scheme slots.
'---------------------------------------------------------------------------------------
'
Private Sub ParseThemeXml(strXmlPath As String)

    Dim objXml As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMNode
    Dim aSlots() As String
    Dim lngSlot As Long
    Dim strColor As String

    Set objXml = New MSXML2.DOMDocument60
    objXml.async = False
    objXml.SetProperty "SelectionLanguage", "XPath"
    objXml.SetProperty "SelectionNamespaces", _
        "xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'"

    If Not objXml.Load(strXmlPath) Then Exit Sub

    aSlots = Split(THEME_SLOTS, ",")

    For lngSlot = 0 To 11
        ' Try srgbClr first (explicit hex), then sysClr (system color with lastClr)
        Set objNode = objXml.SelectSingleNode( _
            "//a:clrScheme/a:" & aSlots(lngSlot) & "/a:srgbClr/@val")
        If Not objNode Is Nothing Then
            this.aBaseColors(lngSlot) = objNode.Text
        Else
            Set objNode = objXml.SelectSingleNode( _
                "//a:clrScheme/a:" & aSlots(lngSlot) & "/a:sysClr/@lastClr")
            If Not objNode Is Nothing Then
                this.aBaseColors(lngSlot) = objNode.Text
            Else
                ' Fallback
                this.aBaseColors(lngSlot) = "808080"
            End If
        End If
    Next lngSlot

    this.blnLoaded = True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsLoaded
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Whether theme colors have been successfully loaded.
'---------------------------------------------------------------------------------------
'
Public Property Get IsLoaded() As Boolean
    IsLoaded = this.blnLoaded
End Property


'---------------------------------------------------------------------------------------
' Procedure : ResolveColor
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Resolve a color from a layout node's properties. Handles literal colors,
'           : theme indices, and tint/shade transforms. Returns "#RRGGBB" or "none".
'           : strPrefix is the property prefix (e.g. "Back", "Fore", "Border").
'---------------------------------------------------------------------------------------
'
Public Function ResolveColor(cNode As clsLayoutNode, strPrefix As String, _
        Optional strDefault As String = "none") As String

    Dim lngColor As Long
    Dim lngThemeIdx As Long
    Dim dblTint As Double
    Dim dblShade As Double
    Dim strBase As String

    ' Check for literal color first
    If cNode.HasProp(strPrefix & "Color") Then
        lngColor = CLng(cNode.PropVal(strPrefix & "Color", 0))
        ' Access stores colors as BGR long; convert to RGB hex
        ResolveColor = "#" & LongToRGBHex(lngColor)
    Else
        ResolveColor = strDefault
    End If

    ' Check for theme color override
    If cNode.HasProp(strPrefix & "ThemeColorIndex") Then
        lngThemeIdx = CLng(cNode.PropVal(strPrefix & "ThemeColorIndex", -1))
        If lngThemeIdx = NO_THEME_INDEX Then
            ' -1 means "no theme"; use literal color above
        ElseIf lngThemeIdx >= 0 And lngThemeIdx <= 11 And this.blnLoaded Then
            strBase = this.aBaseColors(lngThemeIdx)

            ' Apply tint/shade transforms
            dblTint = CDbl(cNode.PropVal(strPrefix & "Tint", 100))
            dblShade = CDbl(cNode.PropVal(strPrefix & "Shade", 100))

            If dblTint < 100 Then
                strBase = ApplyTint(strBase, dblTint / 100#)
            End If
            If dblShade < 100 Then
                strBase = ApplyShade(strBase, dblShade / 100#)
            End If

            ResolveColor = "#" & strBase
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : LongToRGBHex
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Convert an Access BGR long color value to a 6-char RGB hex string.
'---------------------------------------------------------------------------------------
'
Private Function LongToRGBHex(lngColor As Long) As String
    Dim r As Long, g As Long, b As Long
    r = lngColor And &HFF&
    g = (lngColor \ &H100&) And &HFF&
    b = (lngColor \ &H10000) And &HFF&
    LongToRGBHex = Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ApplyTint / ApplyShade
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Apply HSL tint/shade per Microsoft OOXML spec.
'           : Tint: L_new = L + (1 - L) * (1 - tintFactor)
'           : Shade: L_new = L * shadeFactor
'---------------------------------------------------------------------------------------
'
Private Function ApplyTint(strHex As String, dblFactor As Double) As String
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, l As Double
    HexToRGB strHex, r, g, b
    RGBtoHSL r, g, b, h, s, l
    l = l + (1# - l) * (1# - dblFactor)
    HSLtoRGB h, s, l, r, g, b
    ApplyTint = Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
End Function

Private Function ApplyShade(strHex As String, dblFactor As Double) As String
    Dim r As Long, g As Long, b As Long
    Dim h As Double, s As Double, l As Double
    HexToRGB strHex, r, g, b
    RGBtoHSL r, g, b, h, s, l
    l = l * dblFactor
    HSLtoRGB h, s, l, r, g, b
    ApplyShade = Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
End Function


'---------------------------------------------------------------------------------------
' Procedure : HexToRGB / RGBtoHSL / HSLtoRGB
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Color space conversion helpers.
'---------------------------------------------------------------------------------------
'
Private Sub HexToRGB(strHex As String, ByRef r As Long, ByRef g As Long, ByRef b As Long)
    r = CLng("&H" & Mid$(strHex, 1, 2))
    g = CLng("&H" & Mid$(strHex, 3, 2))
    b = CLng("&H" & Mid$(strHex, 5, 2))
End Sub

Private Sub RGBtoHSL(r As Long, g As Long, b As Long, _
        ByRef h As Double, ByRef s As Double, ByRef l As Double)
    Dim dR As Double, dG As Double, dB As Double
    Dim dMax As Double, dMin As Double, dDelta As Double
    dR = r / 255#
    dG = g / 255#
    dB = b / 255#
    dMax = dR: If dG > dMax Then dMax = dG
    If dB > dMax Then dMax = dB
    dMin = dR: If dG < dMin Then dMin = dG
    If dB < dMin Then dMin = dB
    l = (dMax + dMin) / 2#
    If dMax = dMin Then
        h = 0: s = 0
    Else
        dDelta = dMax - dMin
        If l > 0.5 Then s = dDelta / (2# - dMax - dMin) Else s = dDelta / (dMax + dMin)
        If dMax = dR Then
            h = (dG - dB) / dDelta + IIf(dG < dB, 6, 0)
        ElseIf dMax = dG Then
            h = (dB - dR) / dDelta + 2
        Else
            h = (dR - dG) / dDelta + 4
        End If
        h = h / 6#
    End If
End Sub

Private Sub HSLtoRGB(h As Double, s As Double, l As Double, _
        ByRef r As Long, ByRef g As Long, ByRef b As Long)
    Dim Q As Double, p As Double
    If s = 0 Then
        r = CLng(l * 255): g = r: b = r
    Else
        If l < 0.5 Then Q = l * (1# + s) Else Q = l + s - l * s
        p = 2# * l - Q
        r = CLng(HueToRGB(p, Q, h + 1# / 3#) * 255)
        g = CLng(HueToRGB(p, Q, h) * 255)
        b = CLng(HueToRGB(p, Q, h - 1# / 3#) * 255)
    End If
    ' Clamp to valid range
    If r < 0 Then r = 0
    If r > 255 Then r = 255
    If g < 0 Then g = 0
    If g > 255 Then g = 255
    If b < 0 Then b = 0
    If b > 255 Then b = 255
End Sub

Private Function HueToRGB(p As Double, Q As Double, T As Double) As Double
    If T < 0 Then T = T + 1#
    If T > 1 Then T = T - 1#
    If T < 1# / 6# Then
        HueToRGB = p + (Q - p) * 6# * T
    ElseIf T < 0.5 Then
        HueToRGB = Q
    ElseIf T < 2# / 3# Then
        HueToRGB = p + (Q - p) * (2# / 3# - T) * 6#
    Else
        HueToRGB = p
    End If
End Function
