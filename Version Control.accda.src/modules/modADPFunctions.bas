Option Compare Database
Option Explicit



'---------------------------------------------------------------------------------------
' Procedure : ExportADPTriggers
' Author    : awaller
' Date      : 12/14/2016
' Purpose   : Export the triggers
'---------------------------------------------------------------------------------------
'
Public Sub ExportADPTriggers(cOptions As clsOptions, strBaseExportFolder As String)

    Dim colTriggers As New Collection
    Dim rst As ADODB.Recordset
    Dim strSQL As String
    Dim strDef As String
    Dim strFile As String
    Dim varTrg As Variant
    Dim blnFound As Boolean
    Dim dteFileModified As Date
    Dim blnSkip As Boolean
    Dim intObjCnt As Integer
    
    ' Only try this on ADP projects
    If CurrentProject.ProjectType <> acADP Then Exit Sub
    
    ' Build list of triggers in database (from sysobjects)
    strSQL = "SELECT [name],object_name(parent_object_id) AS parent_name, schema_name([schema_id]) AS [schema_name], modify_date FROM sys.objects WHERE type='TR'"
    Set rst = New ADODB.Recordset
    With rst
        .Open strSQL, CurrentProject.Connection, adOpenForwardOnly, adLockReadOnly
        Do While Not .EOF
            strFile = GetSafeFileName(Nz(!schema_name) & "_" & Nz(!Name) & ".sql")
            colTriggers.Add Array(Nz(!Name), Nz(!parent_name), Nz(!schema_name), strFile, Nz(!modify_date))
            .MoveNext
        Loop
        .Close
    End With
    Set rst = Nothing
    
    ' If no triggers, then clear and exit
    If colTriggers.Count = 0 Then
        If FSO.FolderExists(strBaseExportFolder) Then
            ClearTextFilesFromDir strBaseExportFolder, "sql"
            Exit Sub
        End If
    End If
    
    ' Prepare folder
    If Not FSO.FolderExists(strBaseExportFolder) Then VerifyPath strBaseExportFolder
    
    ' Clear all existing files unless we are using fast save.
    If cOptions.UseFastSave Then
    
        ' Loop through saved source files, removing ones that no longer exist in the database.
        strFile = Dir(strBaseExportFolder & "*.sql")
        Do While strFile <> ""
            blnFound = False
            For Each varTrg In colTriggers
                If varTrg(3) = strFile Then
                    ' Found matching object in database
                    blnFound = True
                    Exit For
                End If
            Next varTrg
            If Not blnFound Then
                ' No matching object found
                Kill strBaseExportFolder & strFile
            End If
            strFile = Dir()
        Loop
    Else
        ' Not using fast save.
        ClearTextFilesFromDir strBaseExportFolder, "sql"
    End If
    

    ' Now go through and export the triggers
    For Each varTrg In colTriggers
        
        ' Check for fast save, to see if we can just export the newly changed triggers
        If cOptions.UseFastSave Then
            strFile = strBaseExportFolder & varTrg(3)
            If Not FSO.FileExists(strFile) Then
                blnSkip = False
            Else
                dteFileModified = FileDateTime(strFile)
                If varTrg(4) > dteFileModified Then
                    ' Changed in SQL server
                    blnSkip = False
                Else
                    ' Appears unchanged from the modified dates
                    blnSkip = True
                End If
            End If
        End If
        
        If blnSkip Then
            Log "    (Skipping) [Trigger] - " & varTrg(0), cOptions.ShowDebug
        Else
            ' Export the trigger definition
            strDef = GetSQLObjectDefinitionForADP(varTrg(2) & "." & varTrg(0))
            WriteFile strDef, strBaseExportFolder & varTrg(3)
            ' Show output
            Log "  " & varTrg(0), cOptions.ShowDebug
        End If
        
        ' Increment counter
        intObjCnt = intObjCnt + 1
            
    Next varTrg
    
    ' Display totals
    If cOptions.ShowDebug Then
        Log "[" & intObjCnt & "] triggers exported."
    Else
        Log "[" & intObjCnt & "]"
    End If
    
End Sub