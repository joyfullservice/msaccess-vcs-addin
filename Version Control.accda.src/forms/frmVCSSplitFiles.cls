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
' Purpose   : Start the action to split the files. When a source file is mapped
'           : to multiple destinations, entries are automatically grouped into
'           : batches (one destination per source per batch) and each batch is
'           : processed as a separate git split operation.
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
    Dim dSourceCount As Dictionary
    Dim lngBatchCount As Long
    Dim lngBatch As Long
    Dim lngEntry As Long
    Dim dBatchIndex As Dictionary
    Dim lngIdx As Long
    Dim strBatchPaths() As String
    Dim strBatchNew() As String
    Dim colBatches() As Collection

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
        MsgBox2 T("Validation Failed"), strError, T("Please correct the problem to continue."), vbExclamation
    Else
        ' Proceed with the split after some validation

        ' Get folder from first file (just in case they are from a different location)
        strWorkingFolderBackup = Git.WorkingFolder
        Git.WorkingFolder = FSO.GetParentFolderName(strPaths(0))

        ' Require clean branch with git installation
        If Not Git.IsCleanBranch Then strError = T("Cannot split files in Git when changes are present in the branch")
        If Not Git.Installed Then strError = T("Git must be installed to use this tool.")

        ' Make sure we don't have any errors with the Git commands
        If Len(strError) Then
            MsgBox2 T("Validation Failed"), strError, T("Please correct the problem to continue."), vbExclamation
        Else
            ' Determine how many batches are needed by counting how many times
            ' each source file appears. The max count is the number of batches.
            Set dSourceCount = New Dictionary
            dSourceCount.CompareMode = TextCompare
            lngBatchCount = 1
            For lngEntry = 0 To UBound(strPaths)
                If dSourceCount.Exists(strPaths(lngEntry)) Then
                    dSourceCount(strPaths(lngEntry)) = dSourceCount(strPaths(lngEntry)) + 1
                Else
                    dSourceCount.Add strPaths(lngEntry), 1
                End If
                If dSourceCount(strPaths(lngEntry)) > lngBatchCount Then
                    lngBatchCount = dSourceCount(strPaths(lngEntry))
                End If
            Next lngEntry

            ' Confirm with the user when multiple batches are required
            If lngBatchCount > 1 Then
                If MsgBox2(T("Multiple Batches Required"), _
                    T("Some source files are mapped to multiple destinations. " & _
                    "This requires {0} batches, creating {1} additional commits to preserve history in all files.", _
                    var0:=lngBatchCount, var1:=lngBatchCount * 2), _
                    T("Click OK to proceed or Cancel to abort."), _
                    vbOKCancel + vbInformation) <> vbOK Then
                    GoTo CleanUp
                End If
            End If

            DoCmd.Hourglass True

            If lngBatchCount = 1 Then
                ' Single batch - use the arrays directly
                Git.SplitFilesWithHistory strPaths, strNew, txtCommitMessage
            Else
                ' Distribute entries into batches using round-robin per source file.
                ' dBatchIndex tracks the next batch number for each source.
                Set dBatchIndex = New Dictionary
                dBatchIndex.CompareMode = TextCompare

                ReDim colBatches(0 To lngBatchCount - 1)
                For lngBatch = 0 To lngBatchCount - 1
                    Set colBatches(lngBatch) = New Collection
                Next lngBatch

                For lngEntry = 0 To UBound(strPaths)
                    If dBatchIndex.Exists(strPaths(lngEntry)) Then
                        lngIdx = dBatchIndex(strPaths(lngEntry))
                        dBatchIndex(strPaths(lngEntry)) = lngIdx + 1
                    Else
                        lngIdx = 0
                        dBatchIndex.Add strPaths(lngEntry), 1
                    End If
                    colBatches(lngIdx).Add lngEntry
                Next lngEntry

                ' Execute each batch
                For lngBatch = 0 To lngBatchCount - 1
                    Erase strBatchPaths
                    Erase strBatchNew
                    For lngIdx = 1 To colBatches(lngBatch).Count
                        lngEntry = colBatches(lngBatch)(lngIdx)
                        AddToArray strBatchPaths, strPaths(lngEntry)
                        AddToArray strBatchNew, strNew(lngEntry)
                    Next lngIdx
                    Git.SplitFilesWithHistory strBatchPaths, strBatchNew, txtCommitMessage
                Next lngBatch
            End If

            DoCmd.Hourglass False

            ' Show success message
            MsgBox2 T("Finished"), T("The operation is complete."), _
                T("For additional details, please see `git.log` in the source folder."), vbInformation

            ' Clear existing list
            txtFileList = vbNullString
        End If

CleanUp:
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


'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : Adam Waller
' Date      : 5/9/2025
' Purpose   : Finish up the operation.
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)
    If Not Cancel Then Operation.Finish
End Sub
