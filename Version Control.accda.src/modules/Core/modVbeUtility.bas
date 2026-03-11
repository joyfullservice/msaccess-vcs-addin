Attribute VB_Name = "modVbeUtility"
'---------------------------------------------------------------------------------------
' Module    : modVbeUtility
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : VBA Editor (VBE) operations including code module export, overlay,
'           : compilation, reference management, and add-in project loading.
' Layer     : Core Logic
' Depends on: modObjects, modConstants, modDatabase, modFileAccess, modEncoding,
'           : modErrorHandling, modInstall
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Core")

Private Const ModuleName = "modVbeUtility"


'---------------------------------------------------------------------------------------
' Procedure : ExportVbComponent
' Author    : Adam Waller
' Date      : 5/26/2021
' Purpose   : Export the code module VB component and convert to UTF-8
'---------------------------------------------------------------------------------------
'
Public Sub ExportCodeModule(strName As String, strFile As String)

    Dim strTempFile As String
    Dim strContent As String

    Perf.OperationStart "Export VBE Module"

    ' Export to a temp file so we can convert to UTF-8 encoding
    strTempFile = GetTempFile
    CurrentVBProject.VBComponents(strName).Export strTempFile

    ' Sanitize the VBA code while reading the temp file
    With New clsSourceParser
        .LoadString ReadFile(strTempFile, GetSystemEncoding), edbModule
        strContent = .Sanitize(ectVBA)
    End With

    ' Write the content as UTF-8 to the final destination
    WriteFile strContent, strFile
    DeleteFile strTempFile

    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetFolderAnnotation
' Author    : Adam Waller
' Date      : 3/10/2026
' Purpose   : Return the subfolder path from a Rubberduck-style '@Folder annotation
'           : in a VBE code module. Dots are converted to path separators.
'           : Returns empty string if no annotation found or component has no code.
'           : Example: '@Folder("Core.Utility") returns "Core\Utility\"
'---------------------------------------------------------------------------------------
'
Public Function GetFolderAnnotation(strComponentName As String, Optional strPrefix As String) As String

    Dim cmpItem As VBComponent
    Dim strCode As String
    Dim strUpper As String
    Dim lngPos As Long
    Dim lngStart As Long
    Dim lngEnd As Long

    Const TAG As String = "'@FOLDER("

    LogUnhandledErrors
    On Error Resume Next

    ' Attempt to locate the component in the VBE
    Set cmpItem = CurrentVBProject.VBComponents(strPrefix & strComponentName)
    If cmpItem Is Nothing Then GoTo CleanUp
    If cmpItem.CodeModule Is Nothing Then GoTo CleanUp
    If cmpItem.CodeModule.CountOfLines = 0 Then GoTo CleanUp

    ' Read all code in a single COM call and prepend vbCrLf so that
    ' a line-1 annotation is found by the same pattern as any other line.
    strCode = vbCrLf & cmpItem.CodeModule.Lines(1, 999999)
    strUpper = UCase$(strCode)

    ' Look for '@Folder preceded by a line break and single quote (comment line)
    lngPos = InStr(1, strUpper, vbCrLf & TAG)
    If lngPos = 0 Then Exit Function
    lngPos = lngPos + 2 ' Advance past vbCrLf to the quote character

    ' Warn if a second annotation exists
    If InStr(lngPos + Len(TAG), strUpper, vbCrLf & TAG) > 0 Then
        Log.Add T("WARNING: Multiple @Folder annotations found in {0}. Using first annotation.", _
            var0:=strPrefix & strComponentName), Options.ShowDebug
    End If

    ' Extract the folder path between the double-quote delimiters
    lngStart = InStr(lngPos, strCode, """")
    If lngStart > 0 Then
        lngEnd = InStr(lngStart + 1, strCode, """")
        If lngEnd > lngStart + 1 Then
            GetFolderAnnotation = Replace(Mid$(strCode, lngStart + 1, lngEnd - lngStart - 1), ".", PathSep) & PathSep
        End If
    End If

    CatchAny eelError, "Error reading @Folder annotation for " & strPrefix & strComponentName, ModuleName & ".GetFolderAnnotation"

    Exit Function

CleanUp:
    ' Clear any errors that may cause an early exit
    If Err Then Err.Clear


End Function


'---------------------------------------------------------------------------------------
' Procedure : OverlayCodeModule
' Author    : Adam Waller
' Date      : 10/24/2023
' Purpose   : Overlay VBA code from an object's *.cls file to the form or report
'           : Note that this opens the object in design view, which may slow the build
'           : process if a large number of items are invovled.
'---------------------------------------------------------------------------------------
'
Public Sub OverlayCodeModule(strName As String _
                            , strClassFile As String)

    Const FunctionName As String = ModuleName & ".OverlayCodeModule"

    Dim objModule As VBIDE.CodeModule
    Dim strContent As String
    Dim intType As AcObjectType
    Dim strShortName As String
    Dim cParser As clsSourceParser

    LogUnhandledErrors FunctionName
    'On Error Resume Next

    Set objModule = CurrentVBProject.VBComponents(strName).CodeModule
    If CatchAny(eelError, T("Could not find code module for {0}", var0:=strName), FunctionName) Then Exit Sub

    ' Read class file content
    strContent = ReadFile(strClassFile)
    If strContent = vbNullString Then
        Log.Error eelError, T("Unable to read {0}", var0:=strClassFile), FunctionName
        Exit Sub
    End If

    ' Get object type and short name
    If strName Like "Form_*" Then
        intType = acForm
        strShortName = Mid$(strName, 6)
        DoCmd.OpenForm strShortName, acDesign, , , , acHidden
    ElseIf strName Like "Report_*" Then
        intType = acReport
        strShortName = Mid$(strName, 8)
        DoCmd.OpenReport strShortName, acViewDesign, , , acHidden
    End If

    ' Overlay the VBA code, replacing any existing code.
    Set cParser = New clsSourceParser
    objModule.DeleteLines 1, objModule.CountOfLines
    objModule.AddFromString cParser.StripClassHeader(strContent, False)

    ' Close any form or report object
    Select Case intType
        Case acForm, acReport
            DoCmd.Close intType, strShortName, acSaveYes
    End Select

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RemoveNonBuiltInReferences
' Author    : Adam Waller
' Date      : 10/20/2020
' Purpose   : Remove any references that are not built-in. (Sometimes additional
'           : references are added when creating a new database, but not not really
'           : needed in the completed database when building the project from source.)
'---------------------------------------------------------------------------------------
'
Public Sub RemoveNonBuiltInReferences()

    Dim intCnt As Integer
    Dim strName As String
    Dim ref As Access.Reference

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Perf.OperationStart "Clear References"
    For intCnt = Application.References.Count To 1 Step -1
        Set ref = Application.References(intCnt)
        If Not ref.BuiltIn Then
            strName = ref.Name
            Application.References.Remove ref
            If CatchAny(eelError, "Failed to remove reference: " & strName, _
                ModuleName & ".RemoveNonBuiltInReferences", True, True, True) Then
                Log.Add "  ERROR removing " & strName, False
            Else
                Log.Add "  Removed " & strName, False
            End If
        End If
        Set ref = Nothing
    Next intCnt
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : CompileAndSaveAllModules
' Author    : Adam Waller
' Date      : 7/10/2021
' Purpose   : Compile and save the modules in the current database
'---------------------------------------------------------------------------------------
'
Public Sub CompileAndSaveAllModules()
    Perf.OperationStart "Compile/Save Modules"
    ' Make sure we are running this in the CurrentDB, not the CodeDB
    Set VBE.ActiveVBProject = CurrentVBProject
    DoCmd.RunCommand acCmdCompileAndSaveAllModules
    DoEvents
    Perf.OperationEnd
End Sub


'---------------------------------------------------------------------------------------
' Procedure : PreloadVBE
' Author    : Adam Waller
' Date      : 5/25/2020
' Purpose   : Force Access to load the VBE project. (This can help prevent crashes
'           : when code is run before the VB Project is fully loaded.)
'---------------------------------------------------------------------------------------
'
Public Sub PreloadVBE()
    Dim strName As String
    DoCmd.Hourglass True
    strName = VBE.ActiveVBProject.Name
    DoCmd.Hourglass False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAddInProject
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : Return the VBProject of the MSAccessVCS add-in.
'---------------------------------------------------------------------------------------
'
Public Function GetAddInProject() As VBProject
    Dim oProj As VBProject
    For Each oProj In VBE.VBProjects
        If StrComp(oProj.FileName, GetAddInFileName, vbTextCompare) = 0 Then
            Set GetAddInProject = oProj
            Exit For
        End If
    Next oProj
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadVCSAddIn
' Author    : Adam Waller
' Date      : 11/10/2020
' Purpose   : Load the add-in at the application level so it can stay active
'           : even if the current database is closed.
'           : https://stackoverflow.com/questions/62270088/how-can-i-launch-an-access-add-in-not-com-add-in-from-vba-code
'---------------------------------------------------------------------------------------
'
Public Sub LoadVCSAddIn()
    ' The following lines will load the add-in at the application level,
    ' but will not actually call the function. Ignore the error of function not found.
    LogUnhandledErrors
    On Error Resume Next
    Application.Run GetAddInFileName & "!DummyFunction"
    If Err Then Err.Clear
End Sub
