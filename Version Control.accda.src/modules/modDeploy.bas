Option Compare Database
Option Private Module
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : basDeploy
' Author    : Adam Waller
' Date      : 1/31/2017
' Purpose   : Deploy an update to an Access Database application.
'           : Version number is stored in a custom property in the local database.
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'   USER CONFIGURED OPTIONS
'---------------------------------------------------------------------------------------

' Specify the path to the deployment folder. (UNC path not supported)
Private Const DEPLOY_FOLDER As String = "T:\Apps\Deploy\"

'---------------------------------------------------------------------------------------


' Private variables
Private mVersions As Collection

Private Type udtVersion
    strName As String
    strVersion As String
    dteDate As Date
    strFile As String
End Type


'---------------------------------------------------------------------------------------
' Procedure : Deploy
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Deploys the program for end users to install and run.
'---------------------------------------------------------------------------------------
'
Public Sub Deploy(Optional blnIgnorePendingUpdates As Boolean = False)
    
    Const cstrSpacer As String = "--------------------------------------------------------------"
    Dim strPath As String
    Dim strDeploy As String
    Dim strTools As String
    Dim strName As String
    Dim strCmd As String
    Dim strIcon As String
    
    ' Make sure we don't accidentally deploy a nested library!
    If CodeProject.FullName <> CurrentProject.FullName Then
        Debug.Print " ** WARNING ** " & CodeProject.Name & " is not the top-level project!"
        Debug.Print " Switching to " & CurrentProject.Name & "..."
        Set VBE.ActiveVBProject = GetVBProjectForCurrentDB
        ' Fire off deployment from primary database.
        Run "[" & GetVBProjectForCurrentDB.Name & "].Deploy"
        Exit Sub
    End If
    
    ' Show debug output
    Debug.Print vbCrLf & cstrSpacer
    Debug.Print "Deployment Started - " & Now()
    
    ' Check for any updates to dependent libraries
    If CheckForUpdates Then
        If Not blnIgnorePendingUpdates Then
            Debug.Print cstrSpacer
            Debug.Print " *** UPDATES AVAILABLE *** "
            Debug.Print "Please install before deployment or set flag "
            Debug.Print "to continue deployment anyway. I.e. `Deploy True`" & vbCrLf & cstrSpacer
            Exit Sub
        End If
    End If
    
    ' Check for reference issues with dependent modules
    If HasDuplicateProjects Then
        Select Case Eval("MsgBox('Would you like to run ''LocalizeReferences'' first?@Some VBA projects appear duplicated which usually indicates non-local references.@Select ''No'' to continue anyway or ''Cancel'' to cancel the deployment.@" & _
                "(Library databases that are only used as a part of other applications are typically not deployed as ClickOnce installers.)@',35)")
            Case vbYes
                Call LocalizeReferences
                Exit Sub
            Case vbNo
                ' Continue anyway.
            Case Else
                Exit Sub
        End Select
    End If
    
    ' Increment build number
    IncrementBuildVersion
    
    ' List project and new build number
    Debug.Print " ~ " & VBE.ActiveVBProject.Name & " ~ Version " & AppVersion
    Debug.Print cstrSpacer
    
    ' Update project description
    VBE.ActiveVBProject.Description = "Version " & AppVersion & " deployed on " & Date
    
    ' Check flag for ClickOnce deployment.
    If IsClickOnce Then
        
        ' Get deployment folder (Create if needed)
        strPath = GetDeploymentFolder
        
        ' Copy project files
        Debug.Print "Copying Files";
        Debug.Print vbCrLf & CopyFiles(CodeProject.Path & "\", strPath, True) & " files copied."
        
        ' Get tools folder
        strDeploy = DEPLOY_FOLDER
        strTools = strDeploy & "_Tools\"
        
        ' Copy manifest templates to project
        strName = VBE.ActiveVBProject.Name
        
        ' Build shell command
        strCmd = "cmd /s /c " & strTools & "Deploy.bat """ & strName & """ " & AppVersion
        
        ' Add application icon if one exists in the application folder.
        strIcon = Dir(CodeProject.Path & "\*.ico")
        If strIcon <> "" Then strCmd = strCmd & " """ & strIcon & """"
        
        ' Compile and build clickonce installation
        Shell strCmd, vbNormalFocus
        
        ' Print final status message.
        Debug.Print "Files Copied. Please review command window for any errors." & vbCrLf & cstrSpacer
    Else
        Debug.Print "Version updated." & vbCrLf & cstrSpacer
    End If
    
    ' Update list of latest versions.
    LoadVersionList
    UpdateVersionInList
    SaveVersionList
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetDeploymentFolder
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Returns path to deployment folder, creating if needed.
'---------------------------------------------------------------------------------------
'
Private Function GetDeploymentFolder() As String

    Dim strPath As String
    Dim strProject As String
    Dim strVersion As String
    
    strPath = DEPLOY_FOLDER
    strProject = VBE.ActiveVBProject.Name
    strVersion = AppVersion
    
    ' Build out full path for deployment
    strPath = strPath & strProject
    If Dir(strPath, vbDirectory) = "" Then
        ' Create project folder
        MkDir strPath
    End If
    
    strPath = strPath & "\" & strVersion
    If Dir(strPath, vbDirectory) = "" Then
        ' Create version folder
        MkDir strPath
    End If
    
    ' Return full path
    GetDeploymentFolder = strPath & "\"
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : AppVersion
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Get the version from the database property.
'---------------------------------------------------------------------------------------
'
Public Property Get AppVersion() As String
    Dim strVersion As String
    strVersion = GetDBProperty("AppVersion")
    If strVersion = "" Then strVersion = "1.0.0.0"
    AppVersion = strVersion
End Property


'---------------------------------------------------------------------------------------
' Procedure : AppVersion
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Set version property in current database.
'---------------------------------------------------------------------------------------
'
Public Property Let AppVersion(strVersion As String)
    SetDBProperty "AppVersion", strVersion
End Property


'---------------------------------------------------------------------------------------
' Procedure : GetDBProperty
' Author    : Adam Waller
' Date      : 9/1/2017
' Purpose   : Get a database property
'---------------------------------------------------------------------------------------
'
Public Function GetDBProperty(strName As String) As Variant

    Dim prp As Object   ' Access.AccessObjectProperty
    
    For Each prp In PropertyParent.Properties
        If prp.Name = strName Then
            GetDBProperty = prp.Value
            Exit For
        End If
    Next prp
    
    Set prp = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetDBProperty
' Author    : Adam Waller
' Date      : 9/1/2017
' Purpose   : Set a database property
'---------------------------------------------------------------------------------------
'
Public Sub SetDBProperty(strName As String, varValue, Optional prpType = DB_TEXT)

    Dim prp As Object   ' Access.AccessObjectProperty
    Dim prpAccdb As Property
    Dim blnFound As Boolean
    Dim dbs As Database
    
    For Each prp In PropertyParent.Properties
        If prp.Name = strName Then
            blnFound = True
            ' Skip set on matching value
            If prp.Value = varValue Then Exit Sub
            Exit For
        End If
    Next prp
    
    On Error Resume Next
    If blnFound Then
        PropertyParent.Properties(strName).Value = varValue
    Else
        If CurrentProject.ProjectType = acADP Then
            PropertyParent.Properties.Add strName, varValue
        Else
            ' Normal accdb database property
            Set dbs = CurrentDb
            Set prpAccdb = dbs.CreateProperty(strName, DB_TEXT, varValue)
            dbs.Properties.Append prpAccdb
            Set dbs = Nothing
        End If
    End If
    If Err Then Err.Clear
    On Error GoTo 0

End Sub


'---------------------------------------------------------------------------------------
' Procedure : PropertyParent
' Author    : Adam Waller
' Date      : 1/30/2017
' Purpose   : Get the correct parent type for database properties (including custom)
'---------------------------------------------------------------------------------------
'
Private Function PropertyParent() As Object
    ' Get correct parent project type
    If CurrentProject.ProjectType = acADP Then
        Set PropertyParent = CurrentProject
    Else
        Set PropertyParent = CurrentDb
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : IncrementBuildVersion
' Author    : Adam Waller
' Date      : 1/6/2017
' Purpose   : Increments the build version (1.0.0.x)
'---------------------------------------------------------------------------------------
'
Public Sub IncrementBuildVersion()
    Dim varParts As Variant
    Dim intVer As Integer
    varParts = Split(AppVersion, ".")
    If UBound(varParts) < 3 Then Exit Sub
    intVer = varParts(UBound(varParts))
    varParts(UBound(varParts)) = intVer + 1
    AppVersion = Join(varParts, ".")
End Sub


'---------------------------------------------------------------------------------------
' Procedure : CopyFiles
' Author    : Adam Waller
' Date      : 1/5/2017
' Purpose   : Recursive function to copy files from one folder to another.
'           : (Set to ignore certain files)
'---------------------------------------------------------------------------------------
'
Private Function CopyFiles(strSource As String, strDest As String, blnOverwrite As Boolean) As Double
    
    Dim strFile As String
    Dim dblCnt As Double
    Dim objFSO As New Scripting.FileSystemObject
    Dim oFolder As Scripting.Folder
    Dim oFile As Scripting.File
    Dim blnExists As Boolean
    
    ' Requires FSO to copy open database files. (VBA.FileCopy give a permission denied error.)
    Set oFolder = objFSO.GetFolder(strSource)
    
    ' Copy files then folders
    For Each oFile In objFSO.GetFolder(strSource).Files
        strFile = oFile.Name
        Select Case True
            ' Files to skip
            Case strFile Like ".git*"
            Case strFile Like "*.laccdb"
            Case Else
                blnExists = Dir(strDest & strFile) <> ""
                If blnExists And Not blnOverwrite Then
                    ' Skip this file
                Else
                    If blnExists Then Kill strDest & strFile
                    oFile.Copy strDest & strFile
                    ' Show progress point as each file is copied
                    dblCnt = dblCnt + 1
                    Debug.Print ".";
                End If
        End Select
    Next oFile
    
    ' Copy folders
    For Each oFolder In objFSO.GetFolder(strSource).SubFolders
        strFile = oFolder.Name
        Select Case True
            ' Files to skip
            Case strFile = CodeProject.Name & ".src"    ' This project
            Case strFile Like "*.src"                   ' Other source files
            Case strFile Like ".git*"
            Case Else
                ' Check if folder already exists in destination
                If Dir(strDest & strFile, vbDirectory) = "" Then
                    MkDir strDest & strFile
                    ' Show progress after creating folder but before copying files
                    Debug.Print ".";
                End If
                ' Recursively copy files from this folder
                dblCnt = dblCnt + CopyFiles(strSource & "\" & strFile & "\", strDest & "\" & strFile & "\", blnOverwrite)
        End Select
    Next oFolder
    
    ' Release reference to objects.
    Set objFSO = Nothing
    Set oFile = Nothing
    Set oFolder = Nothing
    
    ' Return count of files copied.
    CopyFiles = dblCnt
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : CheckForUpdates
' Author    : Adam Waller
' Date      : 1/27/2017
' Purpose   : Check for updates to library databases or template modules
'---------------------------------------------------------------------------------------
'
Public Function CheckForUpdates() As Boolean

    Const vbext_rk_Project As Integer = 1
    
    Dim ref As Access.Reference
    Dim varLatest As Variant
    Dim strCurrent As String
    Dim strLatest As String
    Dim objComponent As Object
    Dim strName As String
    Dim intCnt As Integer
    Dim intLines As Integer
    
    ' Reload version file before checking for updates.
    LoadVersionList
    
    ' Check references for updates.
    For Each ref In Application.References
        If ref.Kind = vbext_rk_Project Then
            strCurrent = GetCurrentRefVersion(ref)
            varLatest = GetLatestVersionDetails(ref.Name)
            If IsArray(varLatest) Then
                If UBound(varLatest) > 2 Then
                    strLatest = varLatest(1)
                    If strLatest <> "" Then
                        ' Compare current with latest.
                        If strCurrent <> strLatest Then
                            Debug.Print "UPDATE AVAILABLE: " & ref.Name & " (" & _
                                GetFileNameFromPath(VBE.VBProjects(ref.Name).FileName) & _
                                ") can be updated from " & strCurrent & " to " & strLatest
                            CheckForUpdates = True
                        End If
                    End If
                End If
            End If
        End If
    Next ref
    
    ' Check code modules for updates
    For Each objComponent In GetVBProjectForCurrentDB.VBComponents
        strName = objComponent.Name
        ' Look for matching item in list
        For intCnt = 2 To mVersions.Count
            If UBound(mVersions(intCnt)) = 4 Then
                If (mVersions(intCnt)(0) = strName) _
                    And (mVersions(intCnt)(4) = "Component") Then
                    ' Check for different "version"
                    intLines = GetCodeLineCount(objComponent.CodeModule)
                    If mVersions(intCnt)(1) <> intLines _
                        And mVersions(intCnt)(3) <> CurrentProject.Name Then
                        Debug.Print "MODULE UPDATE AVAILABLE: " & strName & _
                            " can be updated from """ & mVersions(intCnt)(3) & """ (" & _
                            mVersions(intCnt)(1) - intLines & " lines on " & _
                            mVersions(intCnt)(2) & ".)"
                        CheckForUpdates = True
                    End If
                End If
            End If
        Next intCnt
    Next objComponent
    
    Set ref = Nothing
    Set objComponent = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadVersionList
' Author    : Adam Waller
' Date      : 1/27/2017
' Purpose   : Loads a list of the current versions.
'---------------------------------------------------------------------------------------
'
Private Function LoadVersionList() As Boolean
    
    Dim strFile As String
    Dim intFile As Integer
    Dim strLine As String
    
    strFile = DEPLOY_FOLDER & "Latest Versions.csv"
    intFile = FreeFile
    
    ' Initialize collection
    Set mVersions = New Collection
    
    ' Start with header if file does not exist.
    If Dir(strFile) = "" Then
        ' Create a new list.
        mVersions.Add Array("Name", "Version", "Date", "File", "Type")
    Else
        ' Read entries in the file
        Open strFile For Input As #intFile
            Do While Not EOF(intFile)
                Line Input #intFile, strLine
                mVersions.Add Split(strLine, ",")
            Loop
        Close intFile
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : UpdateVersionInList
' Author    : Adam Waller
' Date      : 1/27/2017
' Purpose   : Update the version info in the list of current versions.
'---------------------------------------------------------------------------------------
'
Private Function UpdateVersionInList()
    
    Dim intCnt As Integer
    Dim strName As String
    
    If mVersions Is Nothing Then
        MsgBox "Must load version list first.", vbExclamation
        Exit Function
    End If
    
    ' Structure of list entry:
    'varItem = Array(Name, Version, Date, File, [Type])
    
    ' Get current project name
    strName = GetVBProjectForCurrentDB.Name
    
    ' Look for matching item in list
    For intCnt = 2 To mVersions.Count
        If UBound(mVersions(intCnt)) >= 3 Then
            If mVersions(intCnt)(0) = strName Then
                mVersions.Remove intCnt
                Exit For
            End If
        End If
    Next intCnt
    
    ' Add to list
    mVersions.Add Array(strName, AppVersion, Now, CodeProject.Name, "File"), , , 1
    
    ' Save any code templates
    If CurrentProject.Name = "Code Templates.accdb" Then SaveCodeTemplates
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveCodeTemplates
' Author    : Adam Waller
' Date      : 2/10/2017
' Purpose   : Saves code template modules, using line count as "version".
'---------------------------------------------------------------------------------------
'
Private Sub SaveCodeTemplates()
    
    Dim objComponent As VBIDE.VBComponent
    Dim intCnt As Integer
    Dim blnSkip As Boolean
    Dim intLines As Integer
    Dim strName As String
    
    For Each objComponent In GetVBProjectForCurrentDB.VBComponents
        strName = objComponent.Name
        Select Case strName
        
            ' Skip anything listed here
            Case "basInternal"
                
            ' Any other components
            Case Else
            
                ' Look for matching item in list
                blnSkip = False ' Reset flag
                intLines = GetCodeLineCount(objComponent.CodeModule)
                For intCnt = 2 To mVersions.Count
                    If UBound(mVersions(intCnt)) = 4 Then
                        If (mVersions(intCnt)(0) = strName) _
                            And (mVersions(intCnt)(4) = "Component") Then
                            ' Check for different "version"
                            If mVersions(intCnt)(1) <> intLines Then
                                mVersions.Remove intCnt
                            Else
                                blnSkip = True
                            End If
                            Exit For
                        End If
                    End If
                Next intCnt
                
                ' Add to list
                If Not blnSkip Then mVersions.Add Array(objComponent.Name, intLines, Now, CodeProject.Name, "Component"), , , 1
        End Select
    Next objComponent
    
    Set objComponent = Nothing
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetLatestVersionDetails
' Author    : Adam Waller
' Date      : 1/27/2017
' Purpose   : Return an array of the latest version details.
'---------------------------------------------------------------------------------------
'
Private Function GetLatestVersionDetails(strName As String) As Variant

    Dim varItem As Variant
    
    If mVersions Is Nothing Then LoadVersionList
    
    For Each varItem In mVersions
        If UBound(varItem) > 2 Then
            If varItem(0) = strName Then
                GetLatestVersionDetails = varItem
                Exit Function
            End If
        End If
    Next varItem

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetCurrentRefVersion
' Author    : Adam Waller
' Date      : 1/31/2017
' Purpose   : Return the version of the currently installed reference.
'---------------------------------------------------------------------------------------
'
Private Function GetCurrentRefVersion(ref As Access.Reference) As String

    Dim wrk As Workspace
    Dim dbs As Database
    
    Set wrk = DBEngine(0)
    Set dbs = wrk.OpenDatabase(ref.FullPath, , True)
    
    ' Attempt to read custom property
    On Error Resume Next
    GetCurrentRefVersion = dbs.Properties("AppVersion")
    If Err Then Err.Clear
    On Error GoTo 0
    
    dbs.Close
    DoEvents
    Set dbs = Nothing
    Set wrk = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveVersionList
' Author    : Adam Waller
' Date      : 1/27/2017
' Purpose   : Write the version list to a file.
'---------------------------------------------------------------------------------------
'
Private Function SaveVersionList()
    
    Dim strFile As String
    Dim intFile As Integer
    Dim varLine As Variant
    
    If mVersions Is Nothing Then
        MsgBox "Please load version list before saving", vbExclamation
        Exit Function
    End If
    
    strFile = DEPLOY_FOLDER & "Latest Versions.csv"
    intFile = FreeFile
    
    ' Read entries in the file
    Open strFile For Output As #intFile
        For Each varLine In mVersions
            ' Write in CSV format
            Print #intFile, Join(varLine, ",")
        Next varLine
    Close intFile

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileNameFromPath
' Author    : http://stackoverflow.com/questions/1743328/how-to-extract-file-name-from-path
' Date      : 1/31/2017
' Purpose   : Return file name from path.
'---------------------------------------------------------------------------------------
'
Public Function GetFileNameFromPath(strFullPath As String) As String
    GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetVBProjectForCurrentDB
' Author    : Adam Waller
' Date      : 7/25/2017
' Purpose   : Get the actual VBE project for the current top-level database.
'           : (This is harder than you would think!)
'---------------------------------------------------------------------------------------
'
Private Function GetVBProjectForCurrentDB() As VBProject

    Dim objProj As Object
    Dim strPath As String
    
    strPath = CurrentProject.FullName
    If VBE.ActiveVBProject.FileName = strPath Then
        ' Use currently active project
        Set GetVBProjectForCurrentDB = VBE.ActiveVBProject
    Else
        ' Search for project with matching filename.
        For Each objProj In VBE.VBProjects
            If objProj.FileName = strPath Then
                Set GetVBProjectForCurrentDB = objProj
                Exit For
            End If
        Next objProj
    End If
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetCodeLines
' Author    : Adam Waller
' Date      : 2/14/2017
' Purpose   : A more robust way of counting the lines of code in a module.
'           : (Simply using LineCount can give varying results, due to white
'           :  spacing differences at the end of a code module.)
'---------------------------------------------------------------------------------------
'
Private Function GetCodeLineCount(objCodeModule As Object) As Long
    
    Dim lngLine As Long
    Dim lngLen As Long
    Dim strLine As String
    
    lngLen = objCodeModule.CountOfLines
    
    For lngLine = lngLen To 1 Step -1
        ' Remove line break characters
        strLine = Replace(objCodeModule.Lines(lngLine, lngLine), vbCrLf, "")
        If Trim(strLine) <> "" Then
            ' Found code or comment
            GetCodeLineCount = lngLine
            Exit For
        End If
    Next lngLine
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsClickOnce
' Author    : Adam Waller
' Date      : 2/14/2017
' Purpose   : Returns true if this application should be deployed as a ClickOnce
'           : application. (Stored as custom property rather than module constant
'           : to make updates easier.)
'---------------------------------------------------------------------------------------
'
Private Function IsClickOnce() As Boolean

    Const cstrName As String = "ClickOnce Deployment"
    Dim prp As Object   ' Access.AccessObjectProperty
    Dim strValue As String
    Dim prpAccdb As Object
    Dim dbs As Database
    
    For Each prp In PropertyParent.Properties
        If prp.Name = cstrName Then
            strValue = prp.Value
            Exit For
        End If
    Next prp
    
    Select Case strValue
        
        Case "True", "False"
            ' Use defined value
        
        Case Else
        
            ' Ask user to define preference
            If Eval("MsgBox('Use ClickOnce Deployment for this application?@Select ''Yes'' to create an application " & _
                "that will be installed on the user''s computer, or click ''No'' to simply update the version number.@" & _
                "(Library databases that are only used as a part of other applications are typically not deployed as ClickOnce installers.)@',36)") = vbYes Then
                strValue = "True"
            Else
                strValue = "False"
            End If
            
            ' Save to this database
            If CurrentProject.ProjectType = acADP Then
                PropertyParent.Properties.Add cstrName, strValue
            Else
                ' Normal accdb database property
                Set dbs = CurrentDb
                Set prpAccdb = dbs.CreateProperty(cstrName, DB_TEXT, strValue)
                dbs.Properties.Append prpAccdb
                Set dbs = Nothing
            End If
            
    End Select
    
    Set prp = Nothing
    Set prpAccdb = Nothing
    
    ' Return the existing or newly defined value.
    IsClickOnce = CBool(strValue)
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : LocalizeReferences
' Author    : Adam Waller
' Date      : 2/22/2017
' Purpose   : Make sure references are local
'---------------------------------------------------------------------------------------
'
Public Sub LocalizeReferences()

    Dim oApp As Access.Application
    'Set oApp = New Access.Application
    Set oApp = CreateObject("Access.Application")
    
    With oApp
        .UserControl = True ' Turn visible and stay open.
        .OpenCurrentDatabase DEPLOY_FOLDER & "_Tools\Localize References.accdb"
        .Eval "LocalizeReferencesForRemoteDB(""" & CurrentDb.Name & """)"
    End With
    
    Set oApp = Nothing
    Application.Quit acQuitSaveAll
    
End Sub


'---------------------------------------------------------------------------------------
' Procedure : HasDuplicateProjects
' Author    : Adam Waller
' Date      : 2/22/2017
' Purpose   : Returns true if duplicate projects exist with the same name.
'           : (Typically caused by non-localized references.)
'---------------------------------------------------------------------------------------
'
Private Function HasDuplicateProjects() As Boolean
    
    Dim colProjects As New Collection
    Dim objProj As Object
    Dim strName As String
    Dim varProj As Variant
    
    For Each objProj In VBE.VBProjects
        strName = objProj.Name
        
        ' See if we have already seen this project name.
        For Each varProj In colProjects
            If strName = varProj Then
                HasDuplicateProjects = True
                Exit For
            End If
        Next varProj
        If HasDuplicateProjects Then Exit For
        
        ' Add to list of project names
        colProjects.Add strName
    Next objProj
    
    Set objProj = Nothing
    Set colProjects = Nothing
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : PrintDebugMsg
' Author    : Adam Waller
' Date      : 2/22/2017
' Purpose   : Print a debug message to the immediate window.
'---------------------------------------------------------------------------------------
'
Public Function PrintDebugMsg(strMsg) As String
    Debug.Print strMsg
End Function