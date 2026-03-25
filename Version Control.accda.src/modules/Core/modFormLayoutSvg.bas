Attribute VB_Name = "modFormLayoutSvg"
'---------------------------------------------------------------------------------------
' Module    : modFormLayoutSvg
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Orchestrates form/report layout SVG generation. Reads the sanitized
'           : .form/.report source file, parses into a layout tree, resolves theme
'           : colors, generates SVG, and writes the output file.
' Layer     : Core Logic
' Depends on: clsFormLayoutParser, clsLayoutNode, clsFormLayoutSvgWriter,
'           : clsFormLayoutThemeColors, modFileAccess, modObjects
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Core")

Private Const ModuleName As String = "modFormLayoutSvg"

' Module-level theme colors instance (reused across exports in a single session)
Private m_cThemeColors As clsFormLayoutThemeColors


'---------------------------------------------------------------------------------------
' Procedure : TryExportLayoutSvg
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Generate a layout SVG from a form/report source file if the option is
'           : enabled. Gracefully handles errors without failing the export.
'---------------------------------------------------------------------------------------
'
Public Sub TryExportLayoutSvg(strSourceFile As String, intType As AcObjectType)

    Dim strSvgFile As String
    Dim strContent As String
    Dim strSvg As String
    Dim cParser As clsFormLayoutParser
    Dim cRoot As clsLayoutNode
    Dim cWriter As clsFormLayoutSvgWriter

    ' Only handle forms and reports
    If intType <> acForm And intType <> acReport Then Exit Sub

    ' If the option is disabled, remove any existing SVG and exit
    If Not Options.ExportLayoutSvg Then
        strSvgFile = SwapExtension(strSourceFile, "svg")
        If FSO.FileExists(strSvgFile) Then DeleteFile strSvgFile
        Exit Sub
    End If

    LogUnhandledErrors ModuleName & ".TryExportLayoutSvg"
    On Error Resume Next

    strSvgFile = SwapExtension(strSourceFile, "svg")

    ' Parse the source file
    Perf.OperationStart "Parse Form Layout"
    strContent = ReadSourceFile(strSourceFile)
    Set cParser = New clsFormLayoutParser
    Set cRoot = cParser.Parse(strContent)
    Perf.OperationEnd

    ' Ensure theme colors are loaded (caches after first call)
    EnsureThemeColors

    ' Generate SVG
    Perf.OperationStart "Generate Layout SVG"
    Set cWriter = New clsFormLayoutSvgWriter
    strSvg = cWriter.Generate(cRoot, m_cThemeColors, _
        Options.LayoutSvgScaleMode, Options.LayoutSvgImageEmbed)
    Perf.OperationEnd

    ' Write SVG file (UTF-8 without BOM)
    If Len(strSvg) > 0 Then
        Perf.OperationStart "Write Layout SVG"
        WriteFileNoBom strSvg, strSvgFile
        Perf.OperationEnd
    End If

    CatchAny eelWarning, "Error generating layout SVG for: " & strSourceFile, _
        ModuleName & ".TryExportLayoutSvg", True, True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : EnsureThemeColors
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Load theme colors on first use. Reuses the same instance across
'           : multiple form/report exports in a single export session.
'---------------------------------------------------------------------------------------
'
Private Sub EnsureThemeColors()
    If m_cThemeColors Is Nothing Then
        Perf.OperationStart "Resolve Theme Colors"
        Set m_cThemeColors = New clsFormLayoutThemeColors
        m_cThemeColors.LoadFromThemeFolder Options.GetExportFolder
        Perf.OperationEnd
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ResetThemeColors
' Author    : Adam Waller
' Date      : 03/19/2026
' Purpose   : Reset the cached theme colors (e.g. when starting a new export).
'---------------------------------------------------------------------------------------
'
Public Sub ResetThemeColors()
    Set m_cThemeColors = Nothing
End Sub
