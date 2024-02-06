VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsViewDiff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsViewDiff
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   : Tools for comparing files using 3rd party diff applications
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

' Set this to true to keep the temp files instead
' of automatically cleaning them up at the completion
' of the current operation.
Public PreserveTempFiles As Boolean


' Collection of temp file names (full paths) to clean up later
Private m_colTempFiles As Collection
Private m_strToolName As String


'---------------------------------------------------------------------------------------
' Procedure : DiffStringWithFile
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   : Compare string content with an existing file
'---------------------------------------------------------------------------------------
'
Public Sub StringWithFile(strContent As String, strFile As String)
    Me.Files WriteTemp(strContent), strFile
End Sub


'---------------------------------------------------------------------------------------
' Procedure : FileWithString
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   : Compare with the file on the left and string on the right.
'---------------------------------------------------------------------------------------
'
Public Sub FileWithString(strFile As String, strContent As String)
    Me.Files strFile, WriteTemp(strContent)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : DiffStrings
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   : Show a diff with two strings
'---------------------------------------------------------------------------------------
'
Public Sub Strings(strContent1 As String, strContent2 As String)
    Me.Files WriteTemp(strContent1), WriteTemp(strContent2)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Files
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   : Diff two files
'---------------------------------------------------------------------------------------
'
Public Sub Files(strFile1 As String, strFile2 As String)
    RunCompare strFile1, strFile2
End Sub


'---------------------------------------------------------------------------------------
' Procedure : WriteTemp
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   : Write a string to a temp file, and add to the collection.
'---------------------------------------------------------------------------------------
'
Private Function WriteTemp(strContent As String) As String

    Dim strFile As String

    ' Just in case...
    If m_colTempFiles Is Nothing Then Set m_colTempFiles = New Collection

    ' Create temp file
    strFile = GetTempFile
    WriteFile strContent, strFile
    m_colTempFiles.Add strFile

    ' Return path of temp file
    WriteTemp = strFile

End Function


'---------------------------------------------------------------------------------------
' Procedure : ToolName
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   : Get the name of the Diff tool
'---------------------------------------------------------------------------------------
'
Public Property Get ToolName() As String
    If m_strToolName = vbNullString Then m_strToolName = GetSetting(PROJECT_NAME, "Options", "Diff Tool")
    ToolName = m_strToolName
End Property


'---------------------------------------------------------------------------------------
' Procedure : ToolName
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   : Set the diff tool
'---------------------------------------------------------------------------------------
'
Public Property Let ToolName(strName As String)
    If m_strToolName <> strName Then
        SaveSetting PROJECT_NAME, "Options", "Diff Tool", strName
        m_strToolName = strName
    End If
End Property


'---------------------------------------------------------------------------------------
' Procedure : ClearTempFiles
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   : Delete any remaining temp files after the compare.
'---------------------------------------------------------------------------------------
'
Public Sub ClearTempFiles()

    Dim varFile As Variant

    If Not m_colTempFiles Is Nothing Then
        LogUnhandledErrors
        On Error Resume Next
        For Each varFile In m_colTempFiles
            If FSO.FileExists(varFile) Then DeleteFile CStr(varFile)
        Next varFile
        CatchAny eelNoError, vbNullString
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunCompare
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   : Compare files using the save compare tool.
'---------------------------------------------------------------------------------------
'
Private Sub RunCompare(strFile1 As String, strFile2 As String)

    Dim strCmd As String

    Select Case Me.ToolName

        Case vbNullString
            ' No tool defined
            MsgBox2 "No Diff Tool Specified", _
                "Please specify a diff tool to use for file comparison.", _
                "This setting may be changed in the add-in options.", vbExclamation
            Exit Sub

        Case "WinMerge"
            strCmd = "WinMerge """ & strFile1 & """ """ & strFile2 & """"

        Case "VSCode"
            strCmd = VSCodePath
            If strCmd = vbNullString Then strCmd = "code"
            strCmd = strCmd & " --diff """ & strFile1 & """ """ & strFile2 & """"

        Case Else
            ' Assume they have a custom path defined
            strCmd = Me.ToolName & " """ & strFile1 & """ """ & strFile2 & """"

    End Select

    ' Run command to launch compare
    LogUnhandledErrors
    On Error Resume Next
    If strCmd <> vbNullString Then
        With New WshShell
            .Run strCmd
        End With
    End If
    If Catch(-2147024894) Then Log.Error eelError, "Unable to run '" & Me.ToolName & "' to compare files", ModuleName(Me) & ".RunCompare"
    CatchAny eelError, "Unable to compare files: " & strCmd, ModuleName(Me) & ".RunCompare"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : VSCodePath
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   : Attempt to find the VS Code path from the default install location
'---------------------------------------------------------------------------------------
'
Private Function VSCodePath() As String
    Dim strPath As String
    strPath = Environ("AppData") & "\..\Local\Programs\Microsoft VS Code\Code.exe"
    If FSO.FileExists(strPath) Then VSCodePath = """" & strPath & """"
End Function


'---------------------------------------------------------------------------------------
' Procedure : HasValidCompareTool
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   : Returns true if a compare tool has been selected, and appears to be
'           : valid for use.
'---------------------------------------------------------------------------------------
'
Public Function HasValidCompareTool() As Boolean

    ' Check the current option for compare tool.
    Select Case Me.ToolName

        Case "WinMerge", "VSCode"
            ' These tools probably exist in the environment path
            HasValidCompareTool = True

        Case Else
            ' No tool selected.
    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   : Set up the temp files collection
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Set m_colTempFiles = New Collection
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    If Not Me.PreserveTempFiles Then Me.ClearTempFiles
End Sub
