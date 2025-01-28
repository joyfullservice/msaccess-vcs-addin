Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : cmdAddFormsAndReports_Click
' Author    : Adam Waller
' Date      : 11/9/2023
' Purpose   : Add the forms and reports source files for the project. Doing this
'           : intelligently by only adding items that have a VBA code module.
'---------------------------------------------------------------------------------------
'
Private Sub cmdAddFormsAndReports_Click()

    Dim intType As AcObjectType
    Dim cComponent As IDbComponent
    Dim varKey As Variant
    Dim strFile As String
    Dim strPrefix As String
    Dim cList As clsConcat

    ' Prepare class for new list
    Set cList = New clsConcat
    cList.AppendOnAdd = vbCrLf

    ' Process for forms and reports (2 to 3)
    DoCmd.Hourglass True
    For intType = acForm To acReport

        ' Get component type
        If intType = acForm Then
            Set cComponent = New clsDbForm
            strPrefix = "Form_"
        ElseIf intType = acReport Then
            Set cComponent = New clsDbReport
            strPrefix = "Report_"
        End If

        ' Loop through files
        For Each varKey In cComponent.GetFileList.Keys
            strFile = SwapExtension(CStr(varKey), "cls")
            ' Skip files that already exist
            If Not FSO.FileExists(strFile) Then
                ' Check for code module marker in source file
                If InStr(1, ReadFile(CStr(varKey)), "CodeBehindForm") > 0 Then
                    ' Add to list of files to split
                    cList.Add CStr(varKey), "|", strFile
                End If
            End If
        Next varKey
    Next intType
    DoCmd.Hourglass False
    cmdSplitFiles.SetFocus

    ' See if we found any files to split.
    If cList.Length > 0 Then
        ' Replace existing content.
        txtFileList = cList.GetStr
    Else
        MsgBox2 "No Relevant Files Found", _
            "Could not find any combined form or report source files that contained VBA modules", _
            , vbInformation
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : cmdSplitFiles_Click
' Author    : Adam Waller
' Date      : 5/8/2023
' Purpose   : Start the action to split the files.
'---------------------------------------------------------------------------------------
'
Private Sub cmdSplitFiles_Click()

    Dim varEntries As Variant
    Dim varPaths As Variant
    Dim strPaths() As String
    Dim strNew() As String
    Dim lngLine As Long
    Dim strError As String
    Dim strWorkingFolderBackup As String

    ' Get an array of entries
    varEntries = Split(Nz(txtFileList), vbCrLf)

    ' Loop through lines, building arrays and validating each entry.
    For lngLine = 0 To UBound(varEntries)
        varPaths = Split(varEntries(lngLine), "|")
        If UBound(varPaths) = 1 Then
            ' Perform some validation on the entries
            If Not FSO.FileExists(varPaths(0)) Then strError = "File not found: " & varPaths(0)
            If FSO.FileExists(varPaths(1)) Then strError = "File already exists: " & varPaths(1)
            If varPaths(0) = "c:\example\original.txt" Then strError = "Please use your own file list, not the example."
            If varPaths(0) = varPaths(1) Then strError = "Cannot split to the same file name: " & varPaths(0)
            ' Add to arrays of file paths
            AddToArray strPaths, varPaths(0)
            AddToArray strNew, varPaths(1)
        Else
            If Len(Trim(varEntries(lngLine))) = 0 Then
                ' Ignore blank lines
            Else
                strError = "Expecting two file paths, separated by | character. See line: '" & varPaths(0) & "'"
            End If
        End If
        If Len(strError) Then Exit For
    Next lngLine

    ' Show validation error
    If Len(strError) Then
        MsgBox2 "Validation Failed", strError, "Please correct the problem to continue.", vbExclamation
    Else
        ' Proceed with the split after some validation

        ' Get folder from first file (just in case they are from a different location)
        strWorkingFolderBackup = Git.WorkingFolder
        Git.WorkingFolder = FSO.GetParentFolderName(strPaths(0))

        ' Require clean branch with git installation
        If Not Git.IsCleanBranch Then strError = "Cannot split files in Git when changes are present in the branch"
        If Not Git.Installed Then strError = "Git must be installed to use this tool."

        ' Make sure we don't have any errors with the Git commands
        If Len(strError) Then
            MsgBox2 "Validation Failed", strError, "Please correct the problem to continue.", vbExclamation
        Else
            ' Split the files using git commands
            DoCmd.Hourglass True
            Git.SplitFilesWithHistory strPaths, strNew, txtCommitMessage
            DoCmd.Hourglass False

            ' Show success message
            MsgBox2 "Finished", "The operation is complete.", _
                "For additional details, please see `git.log` in the source folder.", vbInformation

            ' Clear existing list
            txtFileList = vbNullString
        End If

        ' Restore original working folder
        Git.WorkingFolder = strWorkingFolderBackup
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : Adam Waller
' Date      : 3/22/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    Translation.ApplyTo Me
    MakeDialogResizable Me
End Sub
