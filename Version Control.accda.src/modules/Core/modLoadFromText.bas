Attribute VB_Name = "modLoadFromText"
'---------------------------------------------------------------------------------------
' Module    : modLoadFromText
' Author    : bclothier
' Date      : 8/26/2025
' Purpose   : Wraps the Application.LoadFromText method with better error handling
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Core")

Private Type udtThis
    IsInitialized As Boolean
    ErrorFileList As Dictionary
    LastReadError As String
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : Reset
' Author    : bclothier
' Date      : 8/26/2025
' Purpose   : Clear the state
'---------------------------------------------------------------------------------------
'
Public Sub Reset(Optional DeleteErrorFiles As Boolean = False)

    If DeleteErrorFiles And DatabaseFileOpen Then
        Dim strProjectPath As String
        Dim dErrorFiles As Dictionary
        Dim strErrorFilePath As Variant

        strProjectPath = ProjectPath
        If Len(strProjectPath) Then
            Set dErrorFiles = GetFileList(strProjectPath, "errors*.txt")
            For Each strErrorFilePath In dErrorFiles.Keys
                Kill BuildPath2(strProjectPath, strErrorFilePath)
            Next strErrorFilePath
        End If
    End If

    PopulateErrorFileList

End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadFromText
' Author    : bclothier
' Date      : 8/26/2025
' Purpose   : Wraps the Application.LoadFromText to return detailed errors.
'---------------------------------------------------------------------------------------
'
Public Sub LoadFromText(ObjectType As AcObjectType, ObjectName As String, FileName As String)

    Dim ErrNumberOriginal As Long
    Dim ErrSourceOriginal As String
    Dim ErrDescriptionOriginal As String
    Dim ErrHelpFileOriginal As String
    Dim ErrHelpContextOriginal As Long

    LogUnhandledErrors
    On Error GoTo ErrHandler

    If this.IsInitialized = False Then
        PopulateErrorFileList
    End If

    Application.LoadFromText ObjectType, ObjectName, FileName

ExitProc:
    Exit Sub

ErrHandler:
    ' Access error numbers vary by object type (2128 for forms/reports, etc.).
    ' Always check for a newly created errors*.txt file.
    ErrNumberOriginal = Err.Number
    ErrSourceOriginal = Err.Source
    ErrDescriptionOriginal = Err.Description
    ErrHelpFileOriginal = Err.HelpFile
    ErrHelpContextOriginal = Err.HelpContext

    ' Clear Err before ReadErrorFile; FSO and other helpers call LogUnhandledErrors
    ' and would otherwise log this error again as "Unhandled".
    Err.Clear
    ReadErrorFile
    If Len(this.LastReadError) > 0 Then
        ErrDescriptionOriginal = ErrDescriptionOriginal & vbCrLf & this.LastReadError
    End If
    Err.Raise ErrNumberOriginal, ErrSourceOriginal, ErrDescriptionOriginal, ErrHelpFileOriginal, ErrHelpContextOriginal
    Resume ExitProc
    Resume ' for debugging

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PopulateErrorFileList
' Author    : bclothier
' Date      : 8/26/2025
' Purpose   : Initialize the dictionary and check for pre-existing error files.
'---------------------------------------------------------------------------------------
'
Private Sub PopulateErrorFileList()

    Set this.ErrorFileList = New Dictionary
    this.ErrorFileList.CompareMode = TextCompare

    If Not DatabaseFileOpen Then
        this.IsInitialized = False
        Exit Sub
    End If

    Dim strProjectPath As String
    Dim dErrorFiles As Dictionary
    Dim strErrorFilePath As Variant

    strProjectPath = ProjectPath
    If Len(strProjectPath) Then
        Set dErrorFiles = GetFileList(strProjectPath, "errors*.txt")
        For Each strErrorFilePath In dErrorFiles.Keys
            this.ErrorFileList.Add strErrorFilePath, vbNullString
        Next strErrorFilePath
    End If
    this.IsInitialized = True

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ReadErrorFile
' Author    : bclothier
' Date      : 8/26/2025
' Purpose   : Locate the latest error file and read the contents
'---------------------------------------------------------------------------------------
'
Private Function ReadErrorFile() As String

    Dim strProjectPath As String
    Dim dErrorFiles As Dictionary
    Dim strErrorFilePath As Variant
    Dim txtStream As TextStream

    this.LastReadError = vbNullString

    On Error Resume Next

    strProjectPath = ProjectPath
    If Len(strProjectPath) Then
        Set dErrorFiles = GetFileList(strProjectPath, "errors*.txt")
        For Each strErrorFilePath In dErrorFiles.Keys
            If Not this.ErrorFileList.Exists(strErrorFilePath) Then
                ' It's a new file, so read it.
                Set txtStream = FSO.OpenTextFile(BuildPath2(strProjectPath, strErrorFilePath), ForReading, False)
                'Skip the first line
                txtStream.ReadLine

                this.LastReadError = Trim$(txtStream.ReadAll)
                txtStream.Close
                ' Add to the list so it'll be ignored next time.
                this.ErrorFileList.Add strErrorFilePath, vbNullString

                ' We can exit early since only one file will be created
                Exit For
            End If
        Next strErrorFilePath
    End If

    Err.Clear
    On Error GoTo 0

End Function
