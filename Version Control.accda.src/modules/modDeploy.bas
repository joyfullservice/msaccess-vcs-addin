'---------------------------------------------------------------------------------------
' Module    : basDeploy
' Author    : Adam Waller
' Date      : 1/31/2017
' Purpose   : Deploy an update to an Access Database application.
'           : Version number is stored in a custom property in the local database.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

'---------------------------------------------------------------------------------------
'   USER CONFIGURED OPTIONS
'---------------------------------------------------------------------------------------

' Specify the path to the deployment folder. (UNC path not supported)
Private Const DEPLOY_FOLDER As String = "T:\Apps\Deploy\"

' Set this to False to only update the version number and project description.
Private Const DEPLOY_CLICKONCE As Boolean = False

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
    
    ' Increment build number
    IncrementBuildVersion
    
    ' List project and new build number
    Debug.Print " ~ " & VBE.ActiveVBProject.Name & " ~ Version " & AppVersion
    Debug.Print cstrSpacer
    
    ' Update project description
    'VBE.ActiveVBProject.Description = "Version " & AppVersion & " deployed on " & Date
    
    ' Check flag for ClickOnce deployment.
    If DEPLOY_CLICKONCE Then
        
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
        
        ' Compile and build clickonce installation
        Shell "cmd /s /c " & strTools & "Deploy.bat """ & strName & """ " & AppVersion, vbNormalFocus
        
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
Public Function GetDeploymentFolder() As String

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
    
    Dim prp As Object   ' Access.AccessObjectProperty
    Dim objParent As Object
    Dim strVersion As String
    
    For Each prp In PropertyParent.Properties
        If prp.Name = "AppVersion" Then
            strVersion = prp.Value
            Exit For
        End If
    Next prp
    
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
    
    Dim prp As Object   ' Access.AccessObjectProperty
    Dim prpAccdb As Property
    Dim blnFound As Boolean
    
    For Each prp In PropertyParent.Properties
        If prp.Name = "AppVersion" Then
            blnFound = True
            Exit For
        End If
    Next prp
    
    If blnFound Then
        PropertyParent.Properties("AppVersion").Value = strVersion
    Else
        If CodeProject.ProjectType = acADP Then
            PropertyParent.Properties.Add "AppVersion", strVersion
        Else
            ' Normal accdb database property
            Set prpAccdb = CodeDb.CreateProperty("AppVersion", DB_TEXT, strVersion)
            CodeDb.Properties.Append prpAccdb
        End If
    End If

End Property


'---------------------------------------------------------------------------------------
' Procedure : PropertyParent
' Author    : Adam Waller
' Date      : 1/30/2017
' Purpose   : Get the correct parent type for database properties (including custom)
'---------------------------------------------------------------------------------------
'
Private Function PropertyParent() As Object
    ' Get correct parent project type
    If CodeProject.ProjectType = acADP Then
        Set PropertyParent = CurrentProject
    Else
        Set PropertyParent = CodeDb
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
Public Function CopyFiles(strSource As String, strDest As String, blnOverwrite As Boolean) As Double
    
    Dim strFile As String
    Dim dblCnt As Double
    Dim objFSO As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim blnExists As Boolean
    
    ' Requires FSO to copy open database files. (VBA.FileCopy give a permission denied error.)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
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
            Case strFile = CodeProject.Name & ".src"
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
' Purpose   : Check for updates to library databases
'---------------------------------------------------------------------------------------
'
Public Function CheckForUpdates() As Boolean

    Const vbext_rk_Project As Integer = 1
    
    Dim ref As Access.Reference
    Dim varLatest As Variant
    Dim strCurrent As String
    Dim strLatest As String
    
    ' Reload version file before checking for updates.
    LoadVersionList
    
    For Each ref In Application.References
        If ref.Kind = vbext_rk_Project Then
            'strCurrent = GetSession.Utility.GetValue(VBE.ActiveVBProject.Name, "DEPLOY")
            strCurrent = GetCurrentRefVersion(ref)
            varLatest = GetLatestVersionDetails(ref.Name)
            If IsArray(varLatest) Then
                If UBound(varLatest) = 3 Then
                    strLatest = varLatest(1)
                    If strCurrent <> "" Then
                        ' Get current version
                        On Error Resume Next
                        'strLatest = Run("[" & ref.Name & "].AppVersion")
                        If Err Then Err.Clear
                        On Error GoTo 0
                        If strLatest <> "" Then
                            ' Compare current with latest.
                            If strCurrent <> strLatest Then
                                Debug.Print "UPDATE AVAILABLE: " & ref.Name & " (" & GetFileNameFromPath(VBE.VBProjects(ref.Name).fileName) & ") can be updated from " & strCurrent & " to " & strLatest
                                CheckForUpdates = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next ref
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
        mVersions.Add Array("Name", "Version", "Date", "File")
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
    Dim uItem As udtVersion
    Dim blnFound As Boolean
    Dim strName As String
    
    If mVersions Is Nothing Then
        MsgBox "Must load version list first.", vbExclamation
        Exit Function
    End If
    
    ' Structure of list entry:
    'varItem = Array(Name, Version, Date, File)
    
    ' Get current project name
    strName = VBE.ActiveVBProject.Name
    
    ' Look for matching item in list
    For intCnt = 2 To mVersions.Count
        If UBound(mVersions(intCnt)) = 3 Then
            If mVersions(intCnt)(0) = strName Then
                mVersions.Remove intCnt
                Exit For
            End If
        End If
    Next intCnt
    
    ' Add to list
    mVersions.Add Array(strName, AppVersion, Now, CodeProject.Name), , , 1
    
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetLatestVersionDetails
' Author    : Adam Waller
' Date      : 1/27/2017
' Purpose   : Return an array of the latest version details.
'---------------------------------------------------------------------------------------
'
Private Function GetLatestVersionDetails(strName As String) As Variant

    Dim varItem As Variant
    Dim blnFound As Variant
    
    If mVersions Is Nothing Then LoadVersionList
    
    For Each varItem In mVersions
        If UBound(varItem) = 3 Then
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
    Dim strLine As String
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
Function GetFileNameFromPath(strFullPath As String) As String
    GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function