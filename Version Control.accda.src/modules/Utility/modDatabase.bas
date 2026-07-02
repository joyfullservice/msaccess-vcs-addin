Attribute VB_Name = "modDatabase"
'---------------------------------------------------------------------------------------
' Module    : modDatabase
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : General functions for interacting with the current database.
'           : (See modVCSUtility for other functions more specific to this add-in.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Utility")

Private Const ModuleName As String = "modDatabase"

' UDTs for reinterpreting a Long bit pattern as IEEE 754 Single (used by LongToSingle)
Private Type typLong
    Value As Long
End Type

Private Type typSingle
    Value As Single
End Type


'---------------------------------------------------------------------------------------
' Procedure : ProjectPath
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Path/Directory of the current database file.
'---------------------------------------------------------------------------------------
'
Public Function ProjectPath() As String
    ProjectPath = CurrentProject.Path
    If Right$(ProjectPath, 1) <> PathSep Then ProjectPath = ProjectPath & PathSep
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDBProperty
' Author    : Adam Waller
' Date      : 5/6/2021
' Purpose   : Get a database property (Default to MDB version)
'---------------------------------------------------------------------------------------
'
Public Function GetDBProperty(strName As String, Optional dbs As DAO.Database) As Variant

    Dim prp As Object ' DAO.Property
    Dim oParent As Object

    ' Check for database reference
    If Not dbs Is Nothing Then
        Set oParent = dbs.Properties
    Else
        If DatabaseFileOpen Then
            ' Get parent container for properties
            If CurrentProject.ProjectType = acADP Then
                Set oParent = CurrentProject.Properties
            Else
                If dbs Is Nothing Then Set dbs = CurrentDb
                Set oParent = dbs.Properties
            End If
        Else
            ' No database open
            GetDBProperty = vbNullString
            Exit Function
        End If
    End If

    ' Look for property by name
    For Each prp In oParent
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
Public Sub SetDBProperty(ByVal strName As String, ByVal varValue As Variant, Optional ByVal prpType As Long = dbText, Optional dbs As DAO.Database)

    Dim prp As Object ' DAO.Property
    Dim blnFound As Boolean
    Dim oParent As Object

    ' Properties set differently for databases and ADP projects
    If CurrentProject.ProjectType = acADP Then
        Set oParent = CurrentProject.Properties
    Else
        If dbs Is Nothing Then Set dbs = CurrentDb
        Set oParent = dbs.Properties
    End If

    ' Look for property in collection
    For Each prp In oParent
        If prp.Name = strName Then
            ' Check for matching type
            If Not dbs Is Nothing Then
                If prp.Type <> prpType Then
                    ' Remove so we can add it back in with the correct type.
                    dbs.Properties.Delete strName
                    Exit For
                End If
            End If
            blnFound = True
            ' Skip set on matching value
            If prp.Value = varValue Then
                Set dbs = Nothing
            Else
                ' Update value
                prp.Value = varValue
            End If
            Exit Sub
        End If
    Next prp

    ' Add new property
    If Not blnFound Then
        If CurrentProject.ProjectType = acADP Then
            CurrentProject.Properties.Add strName, varValue
        Else
            Set prp = dbs.CreateProperty(strName, prpType, varValue)
            dbs.Properties.Append prp
            Set dbs = Nothing
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SetDAOProperty
' Author    : Adam Waller
' Date      : 5/8/2020
' Purpose   : Updates a DAO property, adding if it does not exist or is the wrong type.
'           : Engine-managed properties (error 3916) are skipped gracefully rather than
'           : aborting the caller. Some properties -- e.g. the FCMin* feature-compatibility
'           : version stamps Access writes when an object uses a newer data type such as
'           : DateTime2 -- can only be set by the database engine. The engine regenerates
'           : these from the object's actual structure, so there is nothing to preserve;
'           : attempting to set them on import raises 3916. We swallow that one error (with
'           : a debug note) and re-raise anything else so real failures still surface to
'           : the caller's handler.
'---------------------------------------------------------------------------------------
'
Public Sub SetDAOProperty(objParent As Object, intType As Integer, strName As String, varValue As Variant)

    Dim prp As DAO.Property
    Dim blnFound As Boolean
    Dim lngErr As Long
    Dim strErrDesc As String

    ' Look through existing properties.
    For Each prp In objParent.Properties
        If prp.Name = strName Then
            blnFound = True
            Exit For
        End If
    Next prp

    ' Apply the change, tolerating engine-managed properties (see header).
    On Error Resume Next
    If blnFound Then
        If prp.Type <> intType Then
            objParent.Properties.Delete strName
            blnFound = False
        ElseIf objParent.Properties(strName).Value <> varValue Then
            objParent.Properties(strName).Value = varValue
        End If
    End If
    If Not blnFound Then
        ' Create property, then append to collection
        Set prp = objParent.CreateProperty(strName, intType, varValue)
        objParent.Properties.Append prp
    End If
    lngErr = Err.Number
    strErrDesc = Err.Description
    On Error GoTo 0

    Select Case lngErr
        Case 0
            ' Success
        Case 3916
            ' Engine-managed property -- skip silently (debug note only).
            Log.Add "    " & T("Skipped engine-managed property '{0}'", var0:=strName), Options.ShowDebug
        Case Else
            ' Surface real errors to the caller, preserving prior behavior.
            Err.Raise lngErr, , strErrDesc
    End Select

End Sub


'---------------------------------------------------------------------------------------
' Procedure : IsEngineManagedProperty
' Author    : Adam Waller
' Date      : 06/19/2026
' Purpose   : Return True for object properties managed by the Access database engine,
'           : which cannot be set by code. The FCMin* family (FCMinDesignVer,
'           : FCMinReadVer, FCMinWriteVer) records the minimum Access build required to
'           : design/read/write the object (e.g. "16.0.12600.10000"). These are
'           : environment/version-specific and engine-regenerated, so we strip them from
'           : exported source to avoid churn. (Import tolerance is handled generically by
'           : SetDAOProperty via error 3916; this is purely an export-cleanliness filter.)
'---------------------------------------------------------------------------------------
'
Public Function IsEngineManagedProperty(strName As String) As Boolean
    IsEngineManagedProperty = StartsWith(strName, "FCMin", vbTextCompare)
End Function


'---------------------------------------------------------------------------------------
' Procedure : FilterEngineManagedProps
' Author    : Adam Waller
' Date      : 06/19/2026
' Purpose   : Return a copy of the property dictionary with any engine-managed
'           : properties removed. (See IsEngineManagedProperty.)
'---------------------------------------------------------------------------------------
'
Public Function FilterEngineManagedProps(dProps As Dictionary) As Dictionary
    Dim varKey As Variant
    Set FilterEngineManagedProps = New Dictionary
    FilterEngineManagedProps.CompareMode = dProps.CompareMode
    For Each varKey In dProps.Keys
        If Not IsEngineManagedProperty(CStr(varKey)) Then
            FilterEngineManagedProps.Add varKey, dProps(varKey)
        End If
    Next varKey
End Function


'---------------------------------------------------------------------------------------
' Procedure : SetAccessObjectProperty
' Author    : Adam Waller
' Date      : 10/13/2017
' Purpose   : Sets a custom access object property.
'---------------------------------------------------------------------------------------
'
Public Sub SetAccessObjectProperty(objItem As AccessObject, strProperty As String, strValue As String)
    Dim prp As AccessObjectProperty
    For Each prp In objItem.Properties
        If StrComp(prp.Name, strProperty, vbTextCompare) = 0 Then
            ' Update value of property.
            prp.Value = strValue
            Exit Sub
        End If
    Next prp
    ' Property not found. Create it.
    objItem.Properties.Add strProperty, strValue
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAccessObjectProperty
' Author    : Adam Waller
' Date      : 10/13/2017
' Purpose   : Get the value of a custom access property
'---------------------------------------------------------------------------------------
'
Public Function GetAccessObjectProperty(objItem As AccessObject, strProperty As String, Optional strDefault As String) As Variant
    Dim prp As AccessObjectProperty
    For Each prp In objItem.Properties
        If StrComp(prp.Name, strProperty, vbTextCompare) = 0 Then
            GetAccessObjectProperty = prp.Value
            Exit Function
        End If
    Next prp
    ' Nothing found. Return default
    GetAccessObjectProperty = strDefault
End Function


'---------------------------------------------------------------------------------------
' Procedure : OpenForm2
' Author    : Adam Waller
' Date      : 6/24/2025
' Purpose   : Wrapper function to open a form in the current database instead of
'           : the add-in when called using the correct method.
'---------------------------------------------------------------------------------------
'
Public Function OpenForm2(FormName, Optional View As AcFormView = acNormal, Optional FilterName, _
    Optional WhereCondition, Optional DataMode As AcFormOpenDataMode = acFormPropertySettings, _
    Optional WindowMode As AcWindowMode = acWindowNormal, Optional OpenArgs)
    DoCmd.OpenForm FormName, View, FilterName, WhereCondition, DataMode, WindowMode, OpenArgs
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsLoaded
' Author    : Adam Waller
' Date      : 9/22/2017
' Purpose   : Returns true if the object is loaded and not in design view.
'---------------------------------------------------------------------------------------
'
Public Function IsLoaded(intType As AcObjectType, strName As String, Optional blnAllowDesignView As Boolean = False) As Boolean

    Dim frm As Form
    Dim rpt As Report
    Dim ctl As Control

    If SysCmd(acSysCmdGetObjectState, intType, strName) <> adStateClosed Then
        If blnAllowDesignView Then
            IsLoaded = True
        Else
            Select Case intType
                Case acForm
                    ' Loop through forms collection, since this includes instances
                    ' of add-in forms that cannot be referenced directly by name.
                    For Each frm In Forms
                        If StrComp(frm.Name, strName, vbTextCompare) = 0 Then
                            IsLoaded = frm.CurrentView <> acCurViewDesign
                            Exit For
                        End If
                    Next frm
                Case acReport
                    ' Loop through reports, looking for matching name.
                    For Each rpt In Reports
                        If StrComp(rpt.Name, strName, vbTextCompare) = 0 Then
                            IsLoaded = rpt.CurrentView <> acCurViewDesign
                            Exit For
                        End If
                    Next rpt
                Case acServerView
                    IsLoaded = CurrentData.AllViews(strName).CurrentView <> acCurViewDesign
                Case acStoredProcedure
                    IsLoaded = CurrentData.AllStoredProcedures(strName).CurrentView <> acCurViewDesign
                Case Else
                    ' Other unsupported object
                    IsLoaded = True
            End Select
        End If
    Else
        ' Could be loaded as subform
        If intType = acForm Then
            For Each frm In Forms
                For Each ctl In frm.Controls
                    If TypeOf ctl Is SubForm Then
                        If ctl.SourceObject = strName Then
                            IsLoaded = True
                            Exit For
                        End If
                    End If
                Next ctl
                If IsLoaded Then Exit For
            Next frm
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : CloseDatabaseObjects
' Author    : Adam Waller
' Date      : 4/1/2022
' Purpose   : Close any open database objects, returns true if no errors were
'           : encountered. (This is run before a major operation like an export or
'           : a merge.) ADP-specific items aren't particularly supported here.
'---------------------------------------------------------------------------------------
'
Public Function CloseDatabaseObjects() As Boolean

    Dim blnSuccess As Boolean
    Dim objItem As AccessObject

    If DebugMode(True) Then On Error GoTo ErrHandler Else On Error GoTo ErrHandler

    Perf.OperationStart "Close Open Objects"

    ' Check forms and reports
    blnSuccess = CloseAllFormsReports

    ' If all forms and reports are closed, proceed with other object types.
    If blnSuccess Then

        ' Macros
        For Each objItem In CurrentProject.AllMacros
            If IsLoaded(acMacro, objItem.Name) Then DoCmd.Close acMacro, objItem.Name
        Next objItem

        ' Tables
        For Each objItem In CurrentData.AllTables
            If IsLoaded(acTable, objItem.Name) Then DoCmd.Close acTable, objItem.Name
        Next objItem

        ' Queries
        For Each objItem In CurrentData.AllQueries
            If IsLoaded(acQuery, objItem.Name) Then DoCmd.Close acQuery, objItem.Name
        Next objItem

    End If

    Perf.OperationEnd
    CloseDatabaseObjects = blnSuccess

    Exit Function

ErrHandler:

    blnSuccess = False

    ' Handle any error message in calling function
    CatchAny eelNoError, T("Unable to close database object"), ModuleName & ".CloseDatabaseObjects", False

End Function


'---------------------------------------------------------------------------------------
' Procedure : CloseAllFormsReports
' Author    : Adam Waller
' Date      : 1/25/2019
' Purpose   : Close all open forms and reports. Returns true if successful.
'---------------------------------------------------------------------------------------
'
Private Function CloseAllFormsReports() As Boolean

    Dim strName As String
    Dim intOpened As Integer
    Dim intItem As Integer

    ' Get count of opened objects
    intOpened = Forms.Count + Reports.Count
    If intOpened > 0 Then
        LogUnhandledErrors
        On Error GoTo ErrorHandler
        ' Loop through forms
        For intItem = Forms.Count - 1 To 0 Step -1
            If Forms(intItem).Caption <> PROJECT_NAME Then
                DoCmd.Close acForm, Forms(intItem).Name
                DoEvents
            End If
            intOpened = intOpened - 1
        Next intItem
        ' Loop through reports
        Do While Reports.Count > 0
            strName = Reports(0).Name
            DoCmd.Close acReport, strName
            DoEvents
            intOpened = intOpened - 1
        Loop
        If intOpened = 0 Then CloseAllFormsReports = True
    Else
        ' No forms or reports currently open.
        CloseAllFormsReports = True
    End If

    Exit Function

ErrorHandler:
    Debug.Print T("Error closing {0}: {1}" & vbCrLf & "{2}", , , , strName, Err.Number, Err.Description)
End Function


'---------------------------------------------------------------------------------------
' Procedure : ObjectExists
' Author    : Adam Waller
' Date      : 3/3/2023
' Purpose   : See if the object exists in the current database/project
'---------------------------------------------------------------------------------------
'
Public Function ObjectExists(intType As AcObjectType, strName As String, Optional blnInCodeDb As Boolean = False) As Boolean

    Dim objTest As Object
    Dim objContainer As Object

    Set objContainer = GetParentContainer(intType, blnInCodeDb)
    If objContainer Is Nothing Then
        Log.Error eelError, T("Parent container not supported for this object type: {0}", var0:=intType), ModuleName & ".ObjectExists"
    Else
        ' Attempt to reference the object by name
        LogUnhandledErrors
        On Error Resume Next
        Set objTest = objContainer(strName)
        ObjectExists = Not Catch(2467)
        If Err Then Err.Clear
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetParentContainer
' Author    : Adam Waller
' Date      : 3/3/2023
' Purpose   : Get the parent container collection for the object type. (Not all types
'           : are supported.)
'           : Also allows you to specify whether to use the CodeDb or CurrentDb
'---------------------------------------------------------------------------------------
'
Public Function GetParentContainer(intType As AcObjectType, Optional blnInCodeDb As Boolean = False) As Object

    Dim objHostData As CurrentData
    Dim objHostProject As CurrentProject

    ' Set the host objects
    If blnInCodeDb Then
        Set objHostData = CodeData
        Set objHostProject = CodeProject
    Else
        Set objHostData = CurrentData
        Set objHostProject = CurrentProject
    End If

    ' Return the associated parent container
    Select Case intType

        ' ADP Specific
        Case acDiagram:             Set GetParentContainer = objHostData.AllDatabaseDiagrams
        Case acFunction:            Set GetParentContainer = objHostData.AllFunctions
        Case acServerView:          Set GetParentContainer = objHostData.AllViews
        Case acStoredProcedure:     Set GetParentContainer = objHostData.AllStoredProcedures

        ' Database objects
        Case acForm:                Set GetParentContainer = objHostProject.AllForms
        Case acMacro:               Set GetParentContainer = objHostProject.AllMacros
        Case acModule:              Set GetParentContainer = objHostProject.AllModules
        Case acQuery:               Set GetParentContainer = objHostData.AllQueries
        Case acReport:              Set GetParentContainer = objHostProject.AllReports
        Case acTable:               Set GetParentContainer = objHostData.AllTables

        ' Unsupported
        Case acTableDataMacro, acDatabaseProperties
        Case Else

    End Select

End Function


'---------------------------------------------------------------------------------------
' Procedure : ProjectIsSelected
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns true if the base project is selected in the VBE
'---------------------------------------------------------------------------------------
'
Public Function ProjectIsSelected() As Boolean
    ProjectIsSelected = (Application.VBE.SelectedVBComponent Is Nothing)
End Function


'---------------------------------------------------------------------------------------
' Procedure : SelectionInActiveProject
' Author    : Adam Waller
' Date      : 5/15/2015
' Purpose   : Returns true if the current selection is in the active project
'---------------------------------------------------------------------------------------
'
Public Function SelectionInActiveProject() As Boolean
    SelectionInActiveProject = (Application.VBE.ActiveVBProject.FileName = GetUncPath(CurrentProject.FullName))
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetCodeVBProject
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Get a reference to the VB Project for the running code.
'---------------------------------------------------------------------------------------
'
Public Function GetCodeVBProject() As VBProject
    Set GetCodeVBProject = GetProjectByName(CodeProject.FullName)
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetProjectByName
' Author    : Adam Waller
' Date      : 5/26/2020
' Purpose   : Return the VBProject by file path. (Also supports network drives)
'---------------------------------------------------------------------------------------
'
Private Function GetProjectByName(ByVal strPath As String) As VBProject

    Dim objProj As VBIDE.VBProject
    Dim strUncPath As String

    ' Use currently active project by default
    Set GetProjectByName = VBE.ActiveVBProject

    ' VBProject filenames are UNC paths
    strUncPath = UCase(GetUncPath(strPath))

    ' Use a guarded FileName read since the collection may contain non-VBA
    ' entries (registered type libraries or wizards) that raise an error
    ' when the .FileName property is accessed. (See GetSafeProjectFileName, #709.)
    If UCase(GetSafeProjectFileName(VBE.ActiveVBProject)) <> strUncPath Then
        ' Search for project with matching filename.
        For Each objProj In VBE.VBProjects
            If UCase(GetSafeProjectFileName(objProj)) = strUncPath Then
                Set GetProjectByName = objProj
                Exit For
            End If
        Next objProj
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : RunInCurrentProject
' Author    : Adam Waller
' Date      : 4/22/2020
' Purpose   : Use the Run command but make sure it is running in the context of the
'           : current project, not the add-in file.
'---------------------------------------------------------------------------------------
'
Public Sub RunSubInCurrentProject(strSubName As String)

    Dim strSub As String
    Dim strCmd As String

    ' Don't need the parentheses after the sub name
    strSub = Replace(strSubName, "()", vbNullString)

    ' Make sure we are not trying to run a function with arguments
    If InStr(strSub, "(") > 0 Then
        MsgBox2 T("Unable to Run Command"), _
            T("Parameters are not supported for this command."), _
            T("If you need to use parameters, please create a wrapper sub or function with" & vbCrLf & _
            "no parameters that you can call instead of {0}.", var0:=strSubName), vbExclamation
        Exit Sub
    End If

    ' Make sure procedure exists in current database
    If Not GlobalProcExists(strSub) Then
        Log.Error eelError, T("The procedure ""{0}"" not found.", var0:=strSub), ModuleName & ".RunSubInCurrentProject"
        Log.Add T("The procedure must be declared as public in a standard module."), False
        Exit Sub
    End If

    ' Build call syntax
    If CurrentVBProject.Name = PROJECT_NAME Then
        ' use full path
        ' Example: Run "c:\full\path\Version Control.SubName"
        With CurrentProject
            strCmd = .Path & PathSep & FSO.GetBaseName(.Name) & "." & strSub
        End With
    Else
        ' use library name
        ' Example: Run "[VBProject].SubName"
        strCmd = "[" & CurrentVBProject.Name & "]." & strSub
    End If

    ' Log any outstanding errors
    LogUnhandledErrors

    ' Stage the current operation, and run the sub
    Operation.Stage
    Perf.OperationStart T("Run {0}", , , , strSub)

    ' Set active VB project to Current DB (not Add-in)
    Set VBE.ActiveVBProject = CurrentVBProject

    Application.Run strCmd
    Perf.OperationEnd
    Operation.Restore

    ' Log any other errors
    CatchAny eelError, T("Error running {0}", , , , strSub), ModuleName & ".RunSubInCurrentProject"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : DatabaseFileOpen
' Author    : Adam Waller
' Date      : 7/14/2020
' Purpose   : Returns true if a database (or ADP project) is currently open.
'---------------------------------------------------------------------------------------
'
Public Function DatabaseFileOpen() As Boolean

    Dim strTest As String

    ' See if we have a reference to a CurrentProject object
    If CurrentProject Is Nothing Then
        DatabaseFileOpen = False
    Else
        ' For ADP projects, CurrentProject may be an invalid object reference
        ' after the database file (adp) is closed.
        LogUnhandledErrors
        On Error Resume Next
        strTest = CurrentProject.FullName
        CatchAny eelNoError, vbNullString
        DatabaseFileOpen = (strTest <> vbNullString)
    End If

    'DatabaseOpen = Workspaces(0).Databases.Count > 0   ' Another approach (Not ADP compatible)

End Function


'---------------------------------------------------------------------------------------
' Procedure : TableExists
' Author    : Adam Waller
' Date      : 3/13/2023
' Purpose   : Returns true if the table object is found in the dabase. (SQL version)
'           : (Includes both local and linked tables, including system tables.)
'---------------------------------------------------------------------------------------
'
Public Function TableExists(strName As String, Optional dbs As DAO.Database) As Boolean

    Dim lngType As Long
    Dim tdf As TableDef

    ' Read type of table from internal system table
    lngType = Nz(DLookup("Type", "MSysObjects", "Name=""" & strName & """ AND Type in (1,4,6)"), 0)

    ' Watch for errors connecting to linked tables
    LogUnhandledErrors
    On Error GoTo ErrHandler

    Select Case lngType
        Case 0  ' Does not exist
            TableExists = False
        Case 1  ' Local table
            TableExists = True
        Case Else
            ' For linked tables, also check availability of target
            If dbs Is Nothing Then Set dbs = CurrentDb
            Set tdf = dbs.TableDefs(strName)
            ' Return true if we find fields in the table
            TableExists = (tdf.Fields.Count > 0)
    End Select

    Exit Function

ErrHandler:
    ' Encountered a problem accessing the table
    TableExists = False

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsLocalTable
' Author    : Adam Waller
' Date      : 3/13/2023
' Purpose   : Returns true if the table exists as a local (not linked) table in the
'           : current database.
'---------------------------------------------------------------------------------------
'
Public Function IsLocalTable(strName As String) As Boolean
    IsLocalTable = Not (DCount("*", "MSysObjects", "Name=""" & strName & """ AND Type = 1") = 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : DeleteObjectIfExists
' Author    : Adam Waller
' Date      : 3/3/2023
' Purpose   : Deletes the object if it exists. This gets really tricky if the object
'           : exists in both the add-in database and the current database. Thankfully
'           : we have a way to work around this for most objects by renaming to a
'           : unique name that doesn't exist in the add-in, and deleting the renamed
'           : object.
'---------------------------------------------------------------------------------------
'
Public Function DeleteObjectIfExists(intType As AcObjectType, strName As String) As Boolean

    Dim blnExistsInAddIn As Boolean
    Dim strTempName As String

    ' If object does not exist in the current database, no need to go further
    If Not ObjectExists(intType, strName) Then Exit Function

    ' Check to see if the object exists in the add-in database. (See note above)
    Select Case intType
        ' Object types used in the add-in
        Case acForm, acMacro, acModule, acQuery, acTable
            blnExistsInAddIn = ObjectExists(intType, strName, True)
    End Select

    ' Trap errors when attempting to delete the object
    LogUnhandledErrors
    On Error Resume Next

    If Not blnExistsInAddIn Then
        ' Nice! We can use a simple call to delete the object
        DoCmd.DeleteObject intType, strName
    Else
        ' This is where it gets fun... If you attempt to delete an object from the
        ' VBA code in the add-in, it will default to operating on the add-in object
        ' first, before looking in the current database. This can cause corruption
        ' in the add-in, in addition to failing to delete the object from the
        ' current database. To work around this, we will rename the object to something
        ' random first, then delete it based on the randomized (non-matching) name.

        ' Build a random temp name that will not collide with the add-in or any existing
        ' object in the current database. (But leave a meaningful clue in the name, in
        ' case something goes wrong and it ends up staying in the database.)
        strTempName = strName & "_DELETE_" & GetStringHash(Perf.MicroTimer)

        ' We need to avoid using DoCmd.Rename for the same reasons
        Select Case intType
            Case acForm
                Log.Error eelError, T("Cannot delete a form with the same name as an add-in form."), _
                    ModuleName & ".DeleteObjectIfExists"
                    Exit Function   ' (Returns False)
            Case acMacro
                ' The rename command seems to work on this object type... (At least
                ' in Access 2010)
                DoCmd.Rename strTempName, intType, strName
            Case acModule
                ' Rename the VBE object
                CurrentVBProject.VBComponents(strName).Name = strTempName
            Case acQuery
                ' Rename the Query Definition object
                CurrentDb.QueryDefs(strName).Name = strTempName
            Case acTable
                ' Rename the Table Definition object
                CurrentDb.TableDefs(strName).Name = strTempName
        End Select

        ' Trap any errors involved in renaming the object
        If Not CatchAny(eelError, T("Error renaming object: {0}", var0:=strName), ModuleName & ".DeleteObjectIfExists") Then
            ' Delete object using the temp name
            DoCmd.DeleteObject intType, strTempName
        End If
    End If

    ' Catch any errors with deleting the object
    CatchAny eelError, T("Error deleting object: {0}", var0:=strName), ModuleName & ".DeleteObjectIfExists"

    ' Return success if the object no longer exists
    DeleteObjectIfExists = Not ObjectExists(intType, strName)

End Function


'---------------------------------------------------------------------------------------
' Procedure : DbVersion
' Author    : Adam Waller
' Date      : 5/4/2021
' Purpose   : Return the database version as an integer. Works in non-English locales
'           : where CInt(CurrentDb.Version) doesn't work correctly.
'---------------------------------------------------------------------------------------
'
Public Function DbVersion() As Integer
    DbVersion = CInt(Split(CurrentDb.Version, ".")(0))
End Function


'---------------------------------------------------------------------------------------
' Procedure : FormLoaded
' Author    : Adam Waller
' Date      : 7/8/2021
' Purpose   : Helps identify if a form has been closed, but is still running code
'           : after the close event.
'---------------------------------------------------------------------------------------
'
Public Function FormLoaded(frmMe As Form) As Boolean
    Dim strName As String
    ' If no forms are open, we already have our answer.  :-)
    If Forms.Count > 0 Then
        ' We will throw an error accessing the name property if the form is closed
        LogUnhandledErrors
        On Error Resume Next
        strName = frmMe.Name
        ' Return true if we were able to read the name property
        FormLoaded = strName <> vbNullString
        If Err Then Err.Clear
    End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : VerifyFocus
' Author    : Adam Waller
' Date      : 7/8/2021
' Purpose   : Verify that a control currently has the focus. (Is the active control)
'---------------------------------------------------------------------------------------
'
Public Function VerifyFocus(ctlWithFocus As Control) As Boolean

    Dim frmParent As Form
    Dim objParent As Object
    Dim ctlCurrentFocus As Control

    ' Determine parent form for control
    Set objParent = ctlWithFocus
    Do While Not TypeOf objParent Is Form
        Set objParent = objParent.Parent
    Loop
    Set frmParent = objParent

    ' Ignore any errors with Screen.* functions
    LogUnhandledErrors
    On Error Resume Next

    ' Verify focus of parent form
    Set frmParent = Screen.ActiveForm
    If Not frmParent Is objParent Then
        Set frmParent = objParent
        frmParent.SetFocus
        DoEvents
    End If

    ' Verify focus of control on form
    Set ctlCurrentFocus = frmParent.ActiveControl
    If Not ctlCurrentFocus Is ctlWithFocus Then
        ctlWithFocus.SetFocus
        DoEvents
    End If

    ' Return true if the control currently has the focus
    VerifyFocus = frmParent.ActiveControl Is ctlWithFocus

    ' Discard any errors
    CatchAny eelNoError, vbNullString, , False

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetAvailableConnectionCount
' Author    : Adam Waller
' Date      : 6/8/2022
' Purpose   : Returns the number of available connections in the current database.
'           : (Access has a maximum of 255 connections that can be made to the current
'           :  database object.) See issue #338
'---------------------------------------------------------------------------------------
'
Public Function GetAvailableConnectionCount()

    Dim colDbs As Collection
    Dim dbs As DAO.Database
    Dim intCnt As Integer

    Set colDbs = New Collection

    LogUnhandledErrors
    On Error Resume Next
    For intCnt = 1 To 300
        Set dbs = CurrentDb
        If Err Then
            ' Probably cannot open any more databases
            Err.Clear
            Exit For
        End If
        colDbs.Add dbs
    Next intCnt

    Set dbs = Nothing

    ' Return count of how many connections we were
    ' able to create before hitting an error
    GetAvailableConnectionCount = intCnt

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsMDE
' Author    : Adam Waller
' Date      : 3/5/2026
' Purpose   : Returns True if the current database is a compiled MDE/ACCDE file.
'           : These files do not contain the original VBA source code.
'---------------------------------------------------------------------------------------
'
Public Function IsMDE() As Boolean
    IsMDE = (GetDBProperty("MDE") = "T")
End Function


'---------------------------------------------------------------------------------------
' Procedure : LongToSingle
' Author    : Adam Waller
' Date      : 3/23/2026
' Purpose   : Reinterprets the raw 4-byte bit pattern of a Long as an IEEE 754
'           : Single-precision float. Used by clsLvPropParser to decode Single-type
'           : properties (e.g. BackTint=100.0 stored as 0x42C80000) from the LvProp blob.
'---------------------------------------------------------------------------------------
'
Public Function LongToSingle(lngVal As Long) As Single
    Dim typLng As typLong
    Dim typSng As typSingle
    typLng.Value = lngVal
    LSet typSng = typLng
    LongToSingle = typSng.Value
End Function
