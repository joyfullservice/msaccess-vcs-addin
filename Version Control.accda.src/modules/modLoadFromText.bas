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

Private Type udtThis
    IsInitialized As Boolean
    FSO As FileSystemObject
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
    If DeleteErrorFiles Then
        Dim strErrorFilePath As String
        Dim strProjectPath

        strProjectPath = ProjectPath
        If Len(strProjectPath) Then
            strErrorFilePath = Dir$(BuildPath2(strProjectPath, "errors*.txt"))
            Do Until Len(strErrorFilePath) = 0
                Kill BuildPath2(strProjectPath, strErrorFilePath)
                strErrorFilePath = Dir()
            Loop
        End If
    End If

    Set this.ErrorFileList = New Dictionary
    this.ErrorFileList.CompareMode = TextCompare
    this.IsInitialized = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : LoadFromText
' Author    : bclothier
' Date      : 8/26/2025
' Purpose   : Wraps the Application.LoadFromText to return detailed errors.
'---------------------------------------------------------------------------------------
'
Public Sub LoadFromText(ObjectType As AcObjectType, ObjectName As String, FileName As String)
    On Error GoTo ErrHandler

    If this.IsInitialized = False Then
        PopulateErrorFileList
    End If

    Application.LoadFromText ObjectType, ObjectName, FileName

ExitProc:
    Exit Sub

ErrHandler:
    If Err.Number = 2128 Then
        Dim ErrNumberOriginal As Long
        Dim ErrSourceOriginal As String
        Dim ErrDescriptionOriginal As String
        Dim ErrHelpFileOriginal As String
        Dim ErrHelpContextOriginal As Long

        ErrNumberOriginal = Err.Number
        ErrSourceOriginal = Err.Source
        ErrDescriptionOriginal = Err.Description
        ErrHelpFileOriginal = Err.HelpFile
        ErrHelpContextOriginal = Err.HelpContext

        ReadErrorFile
        Err.Raise ErrNumberOriginal, ErrSourceOriginal, ErrDescriptionOriginal & this.LastReadError, ErrHelpFileOriginal, ErrHelpContextOriginal
    Else
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
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

    Dim strErrorFilePath As String
    Dim strProjectPath

    strProjectPath = ProjectPath
    If Len(strProjectPath) Then
        strErrorFilePath = Dir$(BuildPath2(strProjectPath, "errors*.txt"))
        Do Until Len(strErrorFilePath) = 0
            this.ErrorFileList.Add strErrorFilePath, vbNullString
            strErrorFilePath = Dir()
        Loop
    End If
    Set this.FSO = FSO
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
    Dim strErrorFilePath As String
    Dim strProjectPath
    Dim txtStream As TextStream

    strProjectPath = ProjectPath
    If Len(strProjectPath) Then
        strErrorFilePath = Dir$(BuildPath2(strProjectPath, "errors*.txt"))
        Do Until Len(strErrorFilePath) = 0
            If Not this.ErrorFileList.Exists(strErrorFilePath) Then
                ' It's a new file, so read it.
                Set txtStream = this.FSO.OpenTextFile(BuildPath2(strProjectPath, strErrorFilePath), ForReading, False)
                'Skip the first line
                txtStream.ReadLine

                this.LastReadError = Trim$(txtStream.ReadAll)
                ' Add to the list so it'll be ignored next time.
                this.ErrorFileList.Add strErrorFilePath, vbNullString

                ' We can exit early since only one file will be created
                Exit Do
            End If
            strErrorFilePath = Dir()
        Loop
    End If
End Function
