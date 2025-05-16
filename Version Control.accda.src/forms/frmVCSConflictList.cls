Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Procedure : txtDiff_Click
' Author    : Adam Waller
' Date      : 5/27/2021
' Purpose   : Launch diff program to review changes.
'           : Note that the changed file is not always the primary source file.
'---------------------------------------------------------------------------------------
'
Private Sub txtDiff_Click()

    Dim strTempFile As String
    Dim strSourceFile As String
    Dim strFileName As String
    Dim cCont As IDbComponent
    Dim dItems As Dictionary
    Dim cItem As IDbComponent

    ' Move focus back to resolution control
    cboResolution.SetFocus
    DoEvents

    ' Make sure we have a valid tool defined
    If Not (modObjects.Diff.HasValidCompareTool) Then
        MsgBox2 "No Compare Tool Defined", _
            "Please specify a compare tool (i.e. WinMerge, VSCode) in the add-in options.", , vbExclamation
        Exit Sub
    End If

    ' Make sure we have a file name to compare
    strFileName = Nz(txtFileName)
    If strFileName = vbNullString Then
        MsgBox2 "File name not found", "A file name is required to compare source files.", , vbExclamation
    Else
        ' Build full path to source file
        strSourceFile = Options.GetExportFolder & strFileName

        ' Check for existing temp file
        strTempFile = VCSIndex.GetTempExportFolder & strFileName
        If Not FSO.FileExists(strTempFile) Then

            ' Has not already been exported. Export a copy that we can use for the compare.
            ' Try to find matching category and file
            For Each cCont In GetContainers(ecfAllObjects)
                If cCont.Category = Nz(txtComponent) Then
                    Set dItems = cCont.GetAllFromDB(False)
                    If cCont.SingleFile Then
                        Set cItem = cCont
                    Else
                        If dItems.Exists(strFileName) Then
                            Set cItem = dItems(strFileName)
                        End If
                    End If
                    ' Build new export file name and export
                    If Not cItem Is Nothing Then cItem.Export strTempFile
                    Exit For
                End If
            Next cCont
        End If

        ' Show comparison if we were able to export a temp file
        If Not FSO.FileExists(strTempFile) Then
            MsgBox2 "Unable to Diff Object", "Unable to produce a temporary diff file with the current database object.", , vbExclamation
        Else
            If Not FSO.FileExists(strSourceFile) Then
                MsgBox2 "Source File Not Found", "Could not find the source file needed to diff this object:", strSourceFile, vbExclamation
            Else
                ' Now that we have both files, diff the files for the user
                If Operation.OperationType = eotExport Then
                    ' Show the database object as the modified version
                    modObjects.Diff.Files strSourceFile, strTempFile
                Else
                    ' Show the source file as the modified version
                    modObjects.Diff.Files strTempFile, strSourceFile
                End If
            End If
        End If
    End If

End Sub
