VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormLayoutSvgWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsFormLayoutSvgWriter
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Generates SVG markup from a parsed clsLayoutNode tree. Uses clsConcat
'           : for efficient string assembly. Walks the tree depth-first in painter's
'           : order (source order) to emit visual elements.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Core")

' Twips to CSS pixels at 96 DPI: px = twips * 96 / 1440 = twips / 15
Private Const TWIPS_PER_PX As Double = 15#

Private Type udtThis
    cSvg As clsConcat
    cTheme As clsFormLayoutThemeColors
    lngIndent As Long
    strScaleMode As String
    strImageEmbed As String
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : Generate
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Generate SVG string from a parsed layout tree.
'---------------------------------------------------------------------------------------
'
Public Function Generate(cRoot As clsLayoutNode, _
        cThemeColors As clsFormLayoutThemeColors, _
        Optional strScaleMode As String = "Universal", _
        Optional strImageEmbed As String = "Small") As String

    Dim cChild As clsLayoutNode
    Dim lngFormWidth As Long
    Dim lngTotalHeight As Long
    Dim dblW As Double
    Dim dblH As Double

    Set this.cTheme = cThemeColors
    this.strScaleMode = strScaleMode
    this.strImageEmbed = strImageEmbed

    Set this.cSvg = New clsConcat
    this.cSvg.AppendOnAdd = vbCrLf
    this.lngIndent = 0

    ' Find the Form or Report root node
    Dim cFormNode As clsLayoutNode
    Set cFormNode = FindFormNode(cRoot)
    If cFormNode Is Nothing Then
        Generate = vbNullString
        Exit Function
    End If

    ' Calculate overall dimensions
    lngFormWidth = CLng(cFormNode.PropVal("Width", 10000))
    lngTotalHeight = CalculateTotalHeight(cFormNode)

    dblW = lngFormWidth / TWIPS_PER_PX
    dblH = lngTotalHeight / TWIPS_PER_PX

    ' SVG header
    AddLine "<!-- generator: msaccess-vcs layout-svg v" & LAYOUT_SVG_GENERATOR_VERSION & _
            ", mode: " & LCase$(strScaleMode) & ", scale: 96dpi -->"
    AddLine "<svg xmlns=""http://www.w3.org/2000/svg"" width=""" & _
            FormatPx(dblW) & """ height=""" & FormatPx(dblH) & _
            """ viewBox=""0 0 " & FormatPx(dblW) & " " & FormatPx(dblH) & _
            """ font-family=""Calibri, sans-serif"">"

    this.lngIndent = this.lngIndent + 1

    ' Collect all Section nodes (they live inside the Defaults block)
    Dim colSections As Collection
    Set colSections = New Collection
    CollectSections cFormNode, colSections

    ' Render sections in order (FormHeader, Detail, FormFooter, etc.)
    Dim dblYOffset As Double
    dblYOffset = 0
    Dim cSection As clsLayoutNode
    For Each cSection In colSections
        RenderSection cSection, dblYOffset, dblW
        dblYOffset = dblYOffset + (cSection.TwipHeight / TWIPS_PER_PX)
    Next cSection

    this.lngIndent = this.lngIndent - 1
    AddLine "</svg>"

    Generate = this.cSvg.GetStr

End Function


'---------------------------------------------------------------------------------------
' Procedure : FindFormNode
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Find the Form or Report node in the parse tree.
'---------------------------------------------------------------------------------------
'
Private Function FindFormNode(cRoot As clsLayoutNode) As clsLayoutNode
    Dim cChild As clsLayoutNode
    For Each cChild In cRoot.Children
        If cChild.ControlType = "Form" Or cChild.ControlType = "Report" Then
            Set FindFormNode = cChild
            Exit Function
        End If
    Next cChild
End Function


'---------------------------------------------------------------------------------------
' Procedure : CalculateTotalHeight
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Sum section heights to get the total form/report height.
'---------------------------------------------------------------------------------------
'
Private Function CalculateTotalHeight(cFormNode As clsLayoutNode) As Long
    Dim colSections As Collection
    Dim cSection As clsLayoutNode
    Dim lngTotal As Long
    Set colSections = New Collection
    CollectSections cFormNode, colSections
    For Each cSection In colSections
        lngTotal = lngTotal + cSection.TwipHeight
    Next cSection
    CalculateTotalHeight = lngTotal
End Function


'---------------------------------------------------------------------------------------
' Procedure : CollectSections
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Recursively find all Section nodes. In SaveAsText format, sections are
'           : nested inside the anonymous "Defaults" block under the Form/Report node.
'---------------------------------------------------------------------------------------
'
Private Sub CollectSections(cNode As clsLayoutNode, colSections As Collection)
    Dim cChild As clsLayoutNode
    For Each cChild In cNode.Children
        If cChild.ControlType = "Section" Then
            colSections.Add cChild
        ElseIf cChild.ControlType = "Defaults" Then
            CollectSections cChild, colSections
        End If
    Next cChild
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderSection
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render a form section (header, detail, footer) and its child controls.
'---------------------------------------------------------------------------------------
'
Private Sub RenderSection(cSection As clsLayoutNode, dblYOffset As Double, dblFormWidth As Double)

    Dim strBackColor As String
    Dim dblH As Double
    Dim strName As String
    Dim cChild As clsLayoutNode

    dblH = cSection.TwipHeight / TWIPS_PER_PX
    strName = cSection.PropVal("Name", "Section")
    strBackColor = ResolveNodeColor(cSection, "Back", "#F0F0F0")

    ' Section group
    AddLine "<g id=""" & EscapeXml(strName) & """>"
    this.lngIndent = this.lngIndent + 1

    ' Section background
    AddLine "<rect x=""0"" y=""" & FormatPx(dblYOffset) & """ width=""" & _
            FormatPx(dblFormWidth) & """ height=""" & FormatPx(dblH) & _
            """ fill=""" & strBackColor & """ />"

    ' Render child controls (may be nested inside a Defaults block)
    RenderChildren cSection, dblYOffset

    this.lngIndent = this.lngIndent - 1
    AddLine "</g>"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderChildren
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render all controls within a node, descending through any Defaults blocks
'           : to find the actual controls. In SaveAsText format, controls within a
'           : section are wrapped in an anonymous "Begin" (Defaults) block.
'---------------------------------------------------------------------------------------
'
Private Sub RenderChildren(cParent As clsLayoutNode, dblSectionY As Double)
    Dim cChild As clsLayoutNode
    For Each cChild In cParent.Children
        If cChild.ControlType = "Defaults" Then
            RenderChildren cChild, dblSectionY
        Else
            RenderControl cChild, dblSectionY
        End If
    Next cChild
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderControl
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render a single control as SVG elements.
'---------------------------------------------------------------------------------------
'
Private Sub RenderControl(cNode As clsLayoutNode, dblSectionY As Double)

    ' Check visibility
    If cNode.PropVal("Visible", vbNullString) = "NotDefault" Then
        ' Control is hidden; render as transparent for now
        ' (future: option to omit entirely)
    End If

    Select Case cNode.ControlType
        Case "Label"
            RenderLabel cNode, dblSectionY
        Case "TextBox"
            RenderTextBox cNode, dblSectionY
        Case "ComboBox"
            RenderComboBox cNode, dblSectionY
        Case "ListBox"
            RenderListBox cNode, dblSectionY
        Case "CommandButton"
            RenderCommandButton cNode, dblSectionY
        Case "Rectangle"
            RenderRectangle cNode, dblSectionY
        Case "Line"
            RenderLine cNode, dblSectionY
        Case "CheckBox"
            RenderCheckBox cNode, dblSectionY
        Case "OptionButton"
            RenderCheckBox cNode, dblSectionY
        Case "ToggleButton"
            RenderCommandButton cNode, dblSectionY
        Case "Image"
            RenderImage cNode, dblSectionY
        Case "Tab"
            RenderTabControl cNode, dblSectionY
        Case "Subform"
            RenderSubform cNode, dblSectionY
        Case "OptionGroup"
            RenderOptionGroup cNode, dblSectionY
        Case "Section"
            ' Nested section (shouldn't normally happen at this level)
        Case "Defaults"
            ' Skip defaults blocks
        Case "Page"
            ' Pages are handled by Tab control
        Case Else
            ' Render unknown controls as a generic rect placeholder
            If cNode.HasProp("Left") And cNode.HasProp("Top") Then
                RenderPlaceholder cNode, dblSectionY
            End If
    End Select

    ' Render attached child controls (e.g. labels associated with checkboxes,
    ' textboxes, comboboxes). Skip types that already handle their own children.
    Select Case cNode.ControlType
        Case "Tab", "OptionGroup", "Section", "Defaults", "Page"
            ' Already handled above
        Case Else
            If cNode.Children.Count > 0 Then
                RenderChildren cNode, dblSectionY
            End If
    End Select

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderLabel
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render a label control with text.
'---------------------------------------------------------------------------------------
'
Private Sub RenderLabel(cNode As clsLayoutNode, dblSectionY As Double)

    Dim dblX As Double, dblY As Double, dblW As Double, dblH As Double
    Dim strCaption As String
    Dim strForeColor As String
    Dim strBackColor As String
    Dim dblFontSize As Double
    Dim strFontWeight As String
    Dim dblOpacity As Double

    GetControlRect cNode, dblSectionY, dblX, dblY, dblW, dblH
    strCaption = CStr(cNode.PropVal("Caption", ""))
    strForeColor = ResolveNodeColor(cNode, "Fore", "#000000")
    dblFontSize = CDbl(cNode.PropVal("FontSize", 11)) * 0.75
    strFontWeight = IIf(CLng(cNode.PropVal("FontWeight", 400)) >= 700, "bold", "normal")
    dblOpacity = IIf(cNode.PropVal("Visible", vbNullString) = "NotDefault", 0.3, 1)

    ' Background (only if BackStyle = 1 = Normal)
    If cNode.PropVal("BackStyle", "0") = "1" Then
        strBackColor = ResolveNodeColor(cNode, "Back", "#FFFFFF")
        AddLine "<rect x=""" & FormatPx(dblX) & """ y=""" & FormatPx(dblY) & _
                """ width=""" & FormatPx(dblW) & """ height=""" & FormatPx(dblH) & _
                """ fill=""" & strBackColor & """" & OpacityAttr(dblOpacity) & " />"
    End If

    ' Text
    If Len(strCaption) > 0 Then
        AddLine "<text x=""" & FormatPx(dblX + 2) & """ y=""" & FormatPx(dblY + dblH * 0.75) & _
                """ font-size=""" & FormatPx(dblFontSize) & """ font-weight=""" & strFontWeight & _
                """ fill=""" & strForeColor & """" & OpacityAttr(dblOpacity) & _
                " clip-path=""inset(0)"">" & EscapeXml(strCaption) & "</text>"
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderTextBox
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render a text box control (input field appearance).
'---------------------------------------------------------------------------------------
'
Private Sub RenderTextBox(cNode As clsLayoutNode, dblSectionY As Double)

    Dim dblX As Double, dblY As Double, dblW As Double, dblH As Double
    Dim strBackColor As String
    Dim strBorderColor As String
    Dim dblOpacity As Double

    GetControlRect cNode, dblSectionY, dblX, dblY, dblW, dblH
    strBackColor = ResolveNodeColor(cNode, "Back", "#FFFFFF")
    strBorderColor = ResolveNodeColor(cNode, "Border", "#ABABAB")
    dblOpacity = IIf(cNode.PropVal("Visible", vbNullString) = "NotDefault", 0.3, 1)

    AddLine "<rect x=""" & FormatPx(dblX) & """ y=""" & FormatPx(dblY) & _
            """ width=""" & FormatPx(dblW) & """ height=""" & FormatPx(dblH) & _
            """ fill=""" & strBackColor & """ stroke=""" & strBorderColor & _
            """ stroke-width=""1""" & OpacityAttr(dblOpacity) & " />"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderComboBox
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render a combo box (text box with dropdown indicator).
'---------------------------------------------------------------------------------------
'
Private Sub RenderComboBox(cNode As clsLayoutNode, dblSectionY As Double)

    Dim dblX As Double, dblY As Double, dblW As Double, dblH As Double
    Dim strBackColor As String
    Dim strBorderColor As String
    Dim dblOpacity As Double
    Dim dblBtnW As Double

    GetControlRect cNode, dblSectionY, dblX, dblY, dblW, dblH
    strBackColor = ResolveNodeColor(cNode, "Back", "#FFFFFF")
    strBorderColor = ResolveNodeColor(cNode, "Border", "#ABABAB")
    dblOpacity = IIf(cNode.PropVal("Visible", vbNullString) = "NotDefault", 0.3, 1)
    dblBtnW = dblH * 0.8

    ' Main box
    AddLine "<rect x=""" & FormatPx(dblX) & """ y=""" & FormatPx(dblY) & _
            """ width=""" & FormatPx(dblW) & """ height=""" & FormatPx(dblH) & _
            """ fill=""" & strBackColor & """ stroke=""" & strBorderColor & _
            """ stroke-width=""1""" & OpacityAttr(dblOpacity) & " />"

    ' Dropdown button area
    AddLine "<line x1=""" & FormatPx(dblX + dblW - dblBtnW) & """ y1=""" & FormatPx(dblY) & _
            """ x2=""" & FormatPx(dblX + dblW - dblBtnW) & """ y2=""" & FormatPx(dblY + dblH) & _
            """ stroke=""" & strBorderColor & """ stroke-width=""1""" & OpacityAttr(dblOpacity) & " />"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderListBox
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render a list box control.
'---------------------------------------------------------------------------------------
'
Private Sub RenderListBox(cNode As clsLayoutNode, dblSectionY As Double)
    RenderTextBox cNode, dblSectionY
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderCommandButton
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render a command button with rounded corners and caption.
'---------------------------------------------------------------------------------------
'
Private Sub RenderCommandButton(cNode As clsLayoutNode, dblSectionY As Double)

    Dim dblX As Double, dblY As Double, dblW As Double, dblH As Double
    Dim strCaption As String
    Dim strBackColor As String
    Dim strForeColor As String
    Dim strBorderColor As String
    Dim dblFontSize As Double
    Dim dblRx As Double
    Dim dblOpacity As Double

    GetControlRect cNode, dblSectionY, dblX, dblY, dblW, dblH
    strCaption = CStr(cNode.PropVal("Caption", ""))
    strBackColor = ResolveNodeColor(cNode, "Back", "#E1E1E1")
    strForeColor = ResolveNodeColor(cNode, "Fore", "#000000")
    strBorderColor = ResolveNodeColor(cNode, "Border", "#ABABAB")
    dblFontSize = CDbl(cNode.PropVal("FontSize", 11)) * 0.75
    dblRx = 3
    dblOpacity = IIf(cNode.PropVal("Visible", vbNullString) = "NotDefault", 0.3, 1)

    ' Button background
    AddLine "<rect x=""" & FormatPx(dblX) & """ y=""" & FormatPx(dblY) & _
            """ width=""" & FormatPx(dblW) & """ height=""" & FormatPx(dblH) & _
            """ rx=""" & FormatPx(dblRx) & """ fill=""" & strBackColor & _
            """ stroke=""" & strBorderColor & """ stroke-width=""1""" & _
            OpacityAttr(dblOpacity) & " />"

    ' Caption text centered
    If Len(strCaption) > 0 Then
        ' Remove leading space and ampersand shortcuts
        strCaption = Replace(strCaption, "&&", "&")
        strCaption = Replace(strCaption, "&", "")
        If Left$(strCaption, 1) = " " Then strCaption = Mid$(strCaption, 2)
        AddLine "<text x=""" & FormatPx(dblX + dblW / 2) & """ y=""" & FormatPx(dblY + dblH / 2 + dblFontSize / 3) & _
                """ font-size=""" & FormatPx(dblFontSize) & """ fill=""" & strForeColor & _
                """ text-anchor=""middle""" & OpacityAttr(dblOpacity) & ">" & EscapeXml(strCaption) & "</text>"
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderRectangle
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render a rectangle shape.
'---------------------------------------------------------------------------------------
'
Private Sub RenderRectangle(cNode As clsLayoutNode, dblSectionY As Double)

    Dim dblX As Double, dblY As Double, dblW As Double, dblH As Double
    Dim strBackColor As String
    Dim strBorderColor As String
    Dim dblOpacity As Double

    GetControlRect cNode, dblSectionY, dblX, dblY, dblW, dblH
    dblOpacity = IIf(cNode.PropVal("Visible", vbNullString) = "NotDefault", 0.3, 1)

    If cNode.PropVal("BackStyle", "0") = "1" Then
        strBackColor = ResolveNodeColor(cNode, "Back", "#FFFFFF")
    Else
        strBackColor = "none"
    End If
    strBorderColor = ResolveNodeColor(cNode, "Border", "none")

    AddLine "<rect x=""" & FormatPx(dblX) & """ y=""" & FormatPx(dblY) & _
            """ width=""" & FormatPx(dblW) & """ height=""" & FormatPx(dblH) & _
            """ fill=""" & strBackColor & """ stroke=""" & strBorderColor & _
            """ stroke-width=""1""" & OpacityAttr(dblOpacity) & " />"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderLine
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render a line control.
'---------------------------------------------------------------------------------------
'
Private Sub RenderLine(cNode As clsLayoutNode, dblSectionY As Double)

    Dim dblX1 As Double, dblY1 As Double, dblX2 As Double, dblY2 As Double
    Dim strColor As String
    Dim dblWidth As Double
    Dim dblOpacity As Double

    dblX1 = cNode.TwipLeft / TWIPS_PER_PX
    dblY1 = dblSectionY + (cNode.TwipTop / TWIPS_PER_PX)
    dblX2 = dblX1 + (cNode.TwipWidth / TWIPS_PER_PX)
    dblY2 = dblY1 + (cNode.TwipHeight / TWIPS_PER_PX)
    strColor = ResolveNodeColor(cNode, "Border", "#000000")
    dblWidth = CDbl(cNode.PropVal("BorderWidth", 1)) / TWIPS_PER_PX
    If dblWidth < 1 Then dblWidth = 1
    dblOpacity = IIf(cNode.PropVal("Visible", vbNullString) = "NotDefault", 0.3, 1)

    AddLine "<line x1=""" & FormatPx(dblX1) & """ y1=""" & FormatPx(dblY1) & _
            """ x2=""" & FormatPx(dblX2) & """ y2=""" & FormatPx(dblY2) & _
            """ stroke=""" & strColor & """ stroke-width=""" & FormatPx(dblWidth) & _
            """" & OpacityAttr(dblOpacity) & " />"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderCheckBox
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render a checkbox or option button as a small box.
'---------------------------------------------------------------------------------------
'
Private Sub RenderCheckBox(cNode As clsLayoutNode, dblSectionY As Double)

    Dim dblX As Double, dblY As Double
    Dim dblSize As Double
    Dim dblOpacity As Double

    dblX = cNode.TwipLeft / TWIPS_PER_PX
    dblY = dblSectionY + (cNode.TwipTop / TWIPS_PER_PX)
    dblSize = 13
    dblOpacity = IIf(cNode.PropVal("Visible", vbNullString) = "NotDefault", 0.3, 1)

    If cNode.ControlType = "OptionButton" Then
        AddLine "<circle cx=""" & FormatPx(dblX + dblSize / 2) & """ cy=""" & _
                FormatPx(dblY + dblSize / 2) & """ r=""" & FormatPx(dblSize / 2) & _
                """ fill=""#FFFFFF"" stroke=""#767676"" stroke-width=""1""" & _
                OpacityAttr(dblOpacity) & " />"
    Else
        AddLine "<rect x=""" & FormatPx(dblX) & """ y=""" & FormatPx(dblY) & _
                """ width=""" & FormatPx(dblSize) & """ height=""" & FormatPx(dblSize) & _
                """ fill=""#FFFFFF"" stroke=""#767676"" stroke-width=""1""" & _
                OpacityAttr(dblOpacity) & " />"
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderImage
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render an image control as a placeholder rect.
'---------------------------------------------------------------------------------------
'
Private Sub RenderImage(cNode As clsLayoutNode, dblSectionY As Double)
    RenderPlaceholder cNode, dblSectionY
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderTabControl
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render a tab control showing the first visible page.
'---------------------------------------------------------------------------------------
'
Private Sub RenderTabControl(cNode As clsLayoutNode, dblSectionY As Double)

    Dim dblX As Double, dblY As Double, dblW As Double, dblH As Double
    Dim strBackColor As String
    Dim cChild As clsLayoutNode
    Dim blnFirstPage As Boolean
    Dim dblOpacity As Double

    GetControlRect cNode, dblSectionY, dblX, dblY, dblW, dblH
    strBackColor = ResolveNodeColor(cNode, "Back", "#F0F0F0")
    dblOpacity = IIf(cNode.PropVal("Visible", vbNullString) = "NotDefault", 0.3, 1)

    ' Tab container
    AddLine "<rect x=""" & FormatPx(dblX) & """ y=""" & FormatPx(dblY) & _
            """ width=""" & FormatPx(dblW) & """ height=""" & FormatPx(dblH) & _
            """ fill=""" & strBackColor & """ stroke=""#ABABAB"" stroke-width=""1""" & _
            OpacityAttr(dblOpacity) & " />"

    ' Render first page's children only
    blnFirstPage = True
    For Each cChild In cNode.Children
        If cChild.ControlType = "Page" Then
            If blnFirstPage Then
                RenderChildren cChild, dblSectionY
                blnFirstPage = False
            End If
        ElseIf cChild.ControlType = "Defaults" Then
            ' Pages may be inside a Defaults block
            Dim cDefaultChild As clsLayoutNode
            For Each cDefaultChild In cChild.Children
                If cDefaultChild.ControlType = "Page" And blnFirstPage Then
                    RenderChildren cDefaultChild, dblSectionY
                    blnFirstPage = False
                End If
            Next cDefaultChild
        End If
    Next cChild

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderSubform
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render a subform as a labeled placeholder box.
'---------------------------------------------------------------------------------------
'
Private Sub RenderSubform(cNode As clsLayoutNode, dblSectionY As Double)

    Dim dblX As Double, dblY As Double, dblW As Double, dblH As Double
    Dim strName As String

    GetControlRect cNode, dblSectionY, dblX, dblY, dblW, dblH
    strName = CStr(cNode.PropVal("Name", "Subform"))

    AddLine "<rect x=""" & FormatPx(dblX) & """ y=""" & FormatPx(dblY) & _
            """ width=""" & FormatPx(dblW) & """ height=""" & FormatPx(dblH) & _
            """ fill=""#F8F8F8"" stroke=""#CCCCCC"" stroke-width=""1"" stroke-dasharray=""4 2"" />"

    AddLine "<text x=""" & FormatPx(dblX + 4) & """ y=""" & FormatPx(dblY + 12) & _
            """ font-size=""9"" fill=""#999999"">[Subform: " & EscapeXml(strName) & "]</text>"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderOptionGroup
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render an option group (frame) and its children.
'---------------------------------------------------------------------------------------
'
Private Sub RenderOptionGroup(cNode As clsLayoutNode, dblSectionY As Double)

    Dim dblX As Double, dblY As Double, dblW As Double, dblH As Double
    Dim strBackColor As String
    Dim cChild As clsLayoutNode
    Dim dblOpacity As Double

    GetControlRect cNode, dblSectionY, dblX, dblY, dblW, dblH
    dblOpacity = IIf(cNode.PropVal("Visible", vbNullString) = "NotDefault", 0.3, 1)

    If cNode.PropVal("BackStyle", "0") = "1" Then
        strBackColor = ResolveNodeColor(cNode, "Back", "#F0F0F0")
    Else
        strBackColor = "none"
    End If

    AddLine "<rect x=""" & FormatPx(dblX) & """ y=""" & FormatPx(dblY) & _
            """ width=""" & FormatPx(dblW) & """ height=""" & FormatPx(dblH) & _
            """ fill=""" & strBackColor & """ stroke=""#ABABAB"" stroke-width=""1""" & _
            OpacityAttr(dblOpacity) & " />"

    ' Render children (may be inside a Defaults block)
    RenderChildren cNode, dblSectionY

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RenderPlaceholder
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Render an unknown or unhandled control as a labeled placeholder.
'---------------------------------------------------------------------------------------
'
Private Sub RenderPlaceholder(cNode As clsLayoutNode, dblSectionY As Double)

    Dim dblX As Double, dblY As Double, dblW As Double, dblH As Double
    Dim strName As String

    GetControlRect cNode, dblSectionY, dblX, dblY, dblW, dblH
    strName = CStr(cNode.PropVal("Name", cNode.ControlType))

    AddLine "<rect x=""" & FormatPx(dblX) & """ y=""" & FormatPx(dblY) & _
            """ width=""" & FormatPx(dblW) & """ height=""" & FormatPx(dblH) & _
            """ fill=""none"" stroke=""#DDDDDD"" stroke-width=""1"" stroke-dasharray=""2 2"" />"

End Sub


'=======================================================================================
' Helper functions
'=======================================================================================


'---------------------------------------------------------------------------------------
' Procedure : GetControlRect
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Get the bounding rectangle for a control in CSS px coordinates.
'---------------------------------------------------------------------------------------
'
Private Sub GetControlRect(cNode As clsLayoutNode, dblSectionY As Double, _
        ByRef dblX As Double, ByRef dblY As Double, _
        ByRef dblW As Double, ByRef dblH As Double)
    dblX = cNode.TwipLeft / TWIPS_PER_PX
    dblY = dblSectionY + (cNode.TwipTop / TWIPS_PER_PX)
    dblW = cNode.TwipWidth / TWIPS_PER_PX
    dblH = cNode.TwipHeight / TWIPS_PER_PX
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResolveNodeColor
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Resolve a color using the theme resolver.
'---------------------------------------------------------------------------------------
'
Private Function ResolveNodeColor(cNode As clsLayoutNode, strPrefix As String, _
        strDefault As String) As String
    If this.cTheme Is Nothing Then
        ResolveNodeColor = strDefault
    Else
        ResolveNodeColor = this.cTheme.ResolveColor(cNode, strPrefix, strDefault)
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddLine
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Add an indented line to the SVG output.
'---------------------------------------------------------------------------------------
'
Private Sub AddLine(strContent As String)
    this.cSvg.Add Space$(this.lngIndent * 2), strContent
End Sub


'---------------------------------------------------------------------------------------
' Procedure : FormatPx
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Format a pixel value as a compact number string.
'---------------------------------------------------------------------------------------
'
Private Function FormatPx(dblValue As Double) As String
    If dblValue = Int(dblValue) Then
        FormatPx = CStr(CLng(dblValue))
    Else
        FormatPx = Format$(dblValue, "0.#")
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : OpacityAttr
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Return an opacity attribute string, or empty if fully opaque.
'---------------------------------------------------------------------------------------
'
Private Function OpacityAttr(dblOpacity As Double) As String
    If dblOpacity < 1 Then
        OpacityAttr = " opacity=""" & Format$(dblOpacity, "0.#") & """"
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : EscapeXml
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Escape special XML characters in text content.
'---------------------------------------------------------------------------------------
'
Private Function EscapeXml(strText As String) As String
    strText = Replace(strText, "&", "&amp;")
    strText = Replace(strText, "<", "&lt;")
    strText = Replace(strText, ">", "&gt;")
    strText = Replace(strText, """", "&quot;")
    EscapeXml = strText
End Function
