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

' Batched table Type cache (table name -> MSysObjects.Type for local/linked tables).
Private m_dTableTypeCache As Dictionary

' Reused temporary query for deterministic table-data XML export (one per operation).
Private Const TABLE_DATA_SORT_QUERY_PREFIX As String = "vcs_tmp_sort_export"
Private m_strTableDataSortQueryName As String

' Temporary staging table used to reconcile table data on a merge (one table at a time).
Private Const TABLE_DATA_STAGING_PREFIX As String = "vcs_tmp_merge_data"

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
' Procedure : CloseOpenObjectsForType
' Author    : Adam Waller
' Date      : 7/17/2026
' Purpose   : Close any open database objects of a single component type. The save
'           : argument controls whether the user is prompted (acSavePrompt) or changes
'           : are auto-saved (acSaveYes). Only UI-openable types are handled; command
'           : bars, properties, references, etc. are no-ops. Module windows are not
'           : closed here — callers should flush unsaved VBA via SaveUnsavedVbaProjectIfNeeded.
'           : Returns False when a close is canceled or fails.
'---------------------------------------------------------------------------------------
'
Public Function CloseOpenObjectsForType(intType As eDatabaseComponentType, intSave As AcCloseSave) As Boolean

    Dim objItem As AccessObject
    Dim intItem As Integer
    Dim intAcType As AcObjectType

    CloseOpenObjectsForType = True

    If DebugMode(True) Then On Error GoTo ErrHandler Else On Error GoTo ErrHandler

    Select Case intType
        Case edbForm
            For intItem = Forms.Count - 1 To 0 Step -1
                If Forms(intItem).Caption <> PROJECT_NAME Then
                    DoCmd.Close acForm, Forms(intItem).Name, intSave
                    DoEvents
                End If
            Next intItem
        Case edbReport
            For intItem = Reports.Count - 1 To 0 Step -1
                DoCmd.Close acReport, Reports(intItem).Name, intSave
                DoEvents
            Next intItem
        Case edbMacro
            intAcType = acMacro
            For Each objItem In CurrentProject.AllMacros
                If SysCmd(acSysCmdGetObjectState, intAcType, objItem.Name) <> adStateClosed Then
                    DoCmd.Close intAcType, objItem.Name, intSave
                End If
            Next objItem
        Case edbQuery
            intAcType = acQuery
            For Each objItem In CurrentData.AllQueries
                If SysCmd(acSysCmdGetObjectState, intAcType, objItem.Name) <> adStateClosed Then
                    DoCmd.Close intAcType, objItem.Name, intSave
                End If
            Next objItem
        Case edbTableDef, edbTableData, edbTableDataMacro
            intAcType = acTable
            For Each objItem In CurrentData.AllTables
                If SysCmd(acSysCmdGetObjectState, intAcType, objItem.Name) <> adStateClosed Then
                    DoCmd.Close intAcType, objItem.Name, intSave
                End If
            Next objItem
    End Select

    Exit Function

ErrHandler:
    CloseOpenObjectsForType = False
    CatchAny eelWarning, T("Error closing open objects"), ModuleName & ".CloseOpenObjectsForType", True, True

End Function


'---------------------------------------------------------------------------------------
' Procedure : CloseOpenObjectsForContainers
' Author    : Adam Waller
' Date      : 7/20/2026
' Purpose   : Close open objects for each container in the collection.
'---------------------------------------------------------------------------------------
'
Public Function CloseOpenObjectsForContainers(colContainers As Collection, intSave As AcCloseSave) As Boolean

    Dim cCategory As IDbComponent

    CloseOpenObjectsForContainers = True

    For Each cCategory In colContainers
        If Not CloseOpenObjectsForType(cCategory.ComponentType, intSave) Then
            CloseOpenObjectsForContainers = False
            Exit Function
        End If
    Next cCategory

End Function


'---------------------------------------------------------------------------------------
' Procedure : SaveUnsavedVbaProjectIfNeeded
' Author    : Adam Waller
' Date      : 7/20/2026
' Purpose   : Save unsaved VBA project changes when a targeted category can contain code.
'---------------------------------------------------------------------------------------
'
Public Sub SaveUnsavedVbaProjectIfNeeded(colContainers As Collection)

    Dim cCategory As IDbComponent
    Dim blnHasVba As Boolean

    For Each cCategory In colContainers
        Select Case cCategory.ComponentType
            Case edbModule, edbForm, edbReport, edbVbeForm
                blnHasVba = True
                Exit For
        End Select
    Next cCategory

    If blnHasVba Then SaveUnsavedVbaProject

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveUnsavedVbaProject
' Author    : Adam Waller
' Date      : 7/29/2026
' Purpose   : Save any unsaved VBA project changes in the current database, so that module
'           : windows do not need to be closed individually. Returns the project's real
'           : Saved state, so False means code is still unsaved right now.
'           :
'           : Delegates to SaveCurrentVBProject. The single-module save this used to perform
'           : does not save the whole project when form and report class modules are dirty,
'           : and it locks the database against other clients — see that procedure.
'           :
'           : Warns when the save did not take, because the caller's next step is to read
'           : code out of the project. Silence here is what made the original bug hard to
'           : find: an export believed it had captured edits that were never written.
'---------------------------------------------------------------------------------------
'
Public Function SaveUnsavedVbaProject() As Boolean
    SaveUnsavedVbaProject = SaveCurrentVBProject
    If Not SaveUnsavedVbaProject Then WarnUnsavedVbaProject
End Function


'---------------------------------------------------------------------------------------
' Procedure : WarnUnsavedVbaProject
' Author    : Adam Waller
' Date      : 8/7/2026
' Purpose   : Tell the user that the VBA project could not be saved automatically, and
'           : what to do about it. Exported source will not include unsaved changes to
'           : form or report class modules, which is silent data loss from the user's
'           : point of view, so this is worth interrupting for.
'           :
'           : Saving the project is the one thing the helper script does that has no
'           : working in-process equivalent (see modVbeUtility.SaveCurrentVBProject for
'           : the mechanisms that were tried and dropped), so a user who turned the
'           : script off (#727) reaches this every time the project is dirty. The message
'           : names that cause when it applies, since the remedy is the same either way
'           : but the reason is not.
'---------------------------------------------------------------------------------------
'
Private Sub WarnUnsavedVbaProject()

    Dim strCause As String
    Dim blnDirty As Boolean

    ' Only claim that changes were missed when the project can be confirmed dirty.
    ' SaveCurrentVBProject also returns False when it could not read the project state
    ' at all, and a warning about unsaved code would be misleading in that case.
    LogUnhandledErrors
    On Error Resume Next
    blnDirty = Not CurrentVBProject.Saved
    If Err Then Err.Clear
    On Error GoTo 0
    If Not blnDirty Then Exit Sub

    If modInstall.UseWorkerScript Then
        strCause = T("The VBA project could not be saved automatically.")
    Else
        strCause = T("The VBA project could not be saved automatically because the " & _
            "helper script is disabled.")
    End If

    Log.Error eelWarning, strCause & " " & _
        T("Unsaved changes to form or report class modules will not be included."), _
        ModuleName & ".SaveUnsavedVbaProject"

    ' A prompt is only useful to somebody who can act on it before reading the results.
    ' Agent and API callers get the log entry instead.
    If Operation.InteractionMode = eimNormal _
        And Operation.Source <> eosMCPTool _
        And Operation.Source <> eosExternalAPI Then
        MsgBox2 T("Unsaved VBA Changes"), strCause, _
            T("Press Save in the Visual Basic Editor, then run this again so the " & _
            "source files include your latest code."), vbExclamation
    End If

End Sub


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
    RunProcInCurrentProject strSubName, , False
End Sub


'---------------------------------------------------------------------------------------
' Procedure : RunProcInCurrentProject
' Author    : Adam Waller
' Date      : 8/12/2026
' Purpose   : Run a procedure in the current project and return whatever it returns.
'           : Identical to RunSubInCurrentProject except that the return value is
'           : surfaced, which is what a validation hook needs. blnRan distinguishes
'           : "ran and returned Empty" from "never ran" (missing, or rejected for
'           : taking parameters) -- a caller that treats a hook failure as a build
'           : failure must not read a refusal as a False return.
'           :
'           : blnWantResult exists so RunSubInCurrentProject keeps calling
'           : Application.Run as a statement. Assigning its result to a Variant
'           : raises "object variable not set" when the procedure returns an object,
'           : which would newly break existing RunAfter* hooks that happen to do so.
'---------------------------------------------------------------------------------------
'
Public Function RunProcInCurrentProject(strSubName As String, _
    Optional ByRef blnRan As Boolean, _
    Optional ByVal blnWantResult As Boolean = True) As Variant

    Dim strSub As String
    Dim strCmd As String
    Dim cPause As clsOperationPause

    blnRan = False

    ' Don't need the parentheses after the sub name
    strSub = Replace(strSubName, "()", vbNullString)

    ' Make sure we are not trying to run a function with arguments
    If InStr(strSub, "(") > 0 Then
        MsgBox2 T("Unable to Run Command"), _
            T("Parameters are not supported for this command."), _
            T("If you need to use parameters, please create a wrapper sub or function with" & vbCrLf & _
            "no parameters that you can call instead of {0}.", var0:=strSubName), vbExclamation
        Exit Function
    End If

    ' Make sure procedure exists in current database
    If Not GlobalProcExists(strSub) Then
        Log.Error eelError, T("The procedure ""{0}"" not found.", var0:=strSub), ModuleName & ".RunSubInCurrentProject"
        Log.Add T("The procedure must be declared as public in a standard module."), False
        Exit Function
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

    ' Pause the current operation while user code runs. The user's procedure is foreign
    ' code: it may raise, open forms, or take its own time, none of which should look
    ' like our operation still running. ResumePause ends the pause on the success path;
    ' Class_Terminate on cPause is the fallback if Application.Run raises.
    Set cPause = Operation.TryPause()
    Perf.OperationStart T("Run {0}", , , , strSub)

    ' Set active VB project to Current DB (not Add-in)
    Set VBE.ActiveVBProject = CurrentVBProject

    If blnWantResult Then
        RunProcInCurrentProject = Application.Run(strCmd)
    Else
        Application.Run strCmd
    End If
    blnRan = True
    Perf.OperationEnd
    If Not cPause Is Nothing Then cPause.ResumePause

    ' Log any other errors
    CatchAny eelError, T("Error running {0}", , , , strSub), ModuleName & ".RunSubInCurrentProject"

End Function


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
' Procedure : GetBigIntRepairFieldNamesFromTableDefXml
' Author    : Adam Waller
' Date      : 7/30/2026
' Purpose   : Return field names whose ExportXML schema cannot represent dbBigInt and
'           : will be mis-created as dbDecimal(38,0) on Application.ImportXML.
'---------------------------------------------------------------------------------------
'
Public Function GetBigIntRepairFieldNamesFromTableDefXml(strXml As String) As Collection

    Const XPATH_CORRUPTED_BIGINT As String = _
        "//*[namespace-uri()='http://www.w3.org/2001/XMLSchema' and local-name()='element'" & _
        " and *[namespace-uri()='http://www.w3.org/2001/XMLSchema' and local-name()='simpleType']" & _
        "/*[namespace-uri()='http://www.w3.org/2001/XMLSchema' and local-name()='restriction' and @base='xsd:decimal']" & _
        "/*[namespace-uri()='http://www.w3.org/2001/XMLSchema' and local-name()='totalDigits' and @value='0']]"

    Dim colNames As New Collection
    Dim objXml As MSXML2.DOMDocument60
    Dim objNodes As MSXML2.IXMLDOMNodeList
    Dim objNode As MSXML2.IXMLDOMNode
    Dim objNameAttr As MSXML2.IXMLDOMNode

    Set GetBigIntRepairFieldNamesFromTableDefXml = colNames
    If Len(strXml) = 0 Then Exit Function

    Set objXml = New MSXML2.DOMDocument60
    objXml.async = False
    If Not objXml.LoadXML(strXml) Then Exit Function

    Set objNodes = objXml.SelectNodes(XPATH_CORRUPTED_BIGINT)
    For Each objNode In objNodes
        Set objNameAttr = objNode.Attributes.getNamedItem("name")
        If Not objNameAttr Is Nothing Then colNames.Add objNameAttr.Text
    Next objNode

    Set GetBigIntRepairFieldNamesFromTableDefXml = colNames

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
' Procedure : GetSystemTableNames
' Author    : Ricardo Hernandez (Notarnet)
' Date      : 8/21/2026
' Purpose   : Return the set of table names that belong to the database engine or to
'           : Access itself, and should therefore be left out of version control.
'           : The system attribute set by the engine is used instead of the object
'           : name, so that user tables which happen to carry the MSys prefix are
'           : exported as regular tables.
'           : dbSystemObject (&H80000002) matches both the engine tables (&H80000000:
'           : MSysObjects, MSysQueries, MSysACEs, MSysRelationships) and the tables
'           : owned by Access (&H00000002: MSysIMEXSpecs, MSysIMEXColumns,
'           : MSysNavPaneGroups, MSysAccessStorage, MSysResources), while user tables
'           : report 0 and linked tables report dbAttachedTable (&H40000000).
'           : Read once per scan so that callers can test names inside their loop
'           : without a lookup per object. Reading the attribute does not open a
'           : linked back-end.
'---------------------------------------------------------------------------------------
'
Public Function GetSystemTableNames() As Dictionary

    Dim dSysTables As Dictionary
    Dim tdf As DAO.TableDef

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    Set dSysTables = New Dictionary
    dSysTables.CompareMode = TextCompare

    For Each tdf In SharedDb.TableDefs
        If (tdf.Attributes And dbSystemObject) <> 0 Then
            dSysTables.Add tdf.Name, True
            ' A user table can carry the system attribute: hiding implementation tables
            ' that way is an old Access idiom. Such a table was exported while the filter
            ' went by name, and is excluded now, so warn instead of dropping it in
            ' silence. This matters because the orphaned-file cleanup deletes its source
            ' file in the same pass.
            If Not (tdf.Name Like "MSys*") Then
                Log.Error eelWarning, T("Table '{0}' carries the system attribute and is " & _
                    "excluded from version control.", var0:=tdf.Name), _
                    ModuleName & ".GetSystemTableNames"
            End If
        End If
    Next tdf

    ' Tables created by the engine that carry no system attribute, and so have to be
    ' listed by name. (MSysCompactError is left behind by a failed compact operation.)
    ' Listing a name that does not exist in this database is harmless.
    If Not dSysTables.Exists("MSysCompactError") Then dSysTables.Add "MSysCompactError", True

    ' A partial set (the handler above covers the whole loop) is detectable afterwards.
    Log.Add "System tables: " & dSysTables.Count, Options.ShowDebug

    Set GetSystemTableNames = dSysTables

    CatchAny eelError, T("Error reading the list of system tables"), ModuleName & ".GetSystemTableNames"

End Function



'---------------------------------------------------------------------------------------
' Procedure : TableIndexesAvailable
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Return true if the index collection is available. Without the error handling
'           : this may throw an error if a linked table is not accessible during export.
'---------------------------------------------------------------------------------------
'
Public Function TableIndexesAvailable(tdf As DAO.TableDef) As Boolean

    Dim lngTest As Long

    LogUnhandledErrors
    On Error Resume Next
    lngTest = tdf.Indexes.Count
    If Err Then
        Err.Clear
    Else
        TableIndexesAvailable = True
    End If
    CatchAny eelNoError, vbNullString, , False

End Function


'---------------------------------------------------------------------------------------
' Function  : GetTableSortFields
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Return the field names and DAO types to use when sorting table data for
'           : deterministic export. Prefers primary key, then unique+required index,
'           : then all non-binary fields.
'---------------------------------------------------------------------------------------
'
Public Function GetTableSortFields(tdf As DAO.TableDef) As Dictionary

    Dim dFields As Dictionary
    Dim fld As DAO.Field

    Set dFields = GetTableMergeKey(tdf)
    If dFields.Count > 0 Then
        Set GetTableSortFields = dFields
        Exit Function
    End If

    For Each fld In tdf.Fields
        If Not IsBinaryTableFieldType(fld.Type) Then dFields.Add fld.Name, fld.Type
    Next fld

    Set GetTableSortFields = dFields

End Function


'---------------------------------------------------------------------------------------
' Function  : GetTableMergeKey
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Return the field names and DAO types of the primary key, or failing that a
'           : unique and required index, for use as a row identity when merging table
'           : data. Returns an empty dictionary when the table has no such index.
'           :
'           : Unlike GetTableSortFields, this never falls back to the full field list.
'           : Sorting only needs a deterministic order, but merging needs each row in the
'           : source file to match at most one row in the table, so a non-unique key
'           : would silently update or delete the wrong rows.
'---------------------------------------------------------------------------------------
'
Public Function GetTableMergeKey(tdf As DAO.TableDef) As Dictionary

    Dim dFields As Dictionary
    Dim idx As DAO.Index
    Dim idxFld As Object

    Set dFields = New Dictionary
    dFields.CompareMode = vbTextCompare
    Set GetTableMergeKey = dFields

    If Not TableIndexesAvailable(tdf) Then Exit Function

    For Each idx In tdf.Indexes
        If idx.Primary Then
            For Each idxFld In idx.Fields
                dFields.Add idxFld.Name, tdf.Fields(idxFld.Name).Type
            Next idxFld
            Exit Function
        End If
    Next idx

    For Each idx In tdf.Indexes
        If idx.Unique And idx.Required Then
            For Each idxFld In idx.Fields
                dFields.Add idxFld.Name, tdf.Fields(idxFld.Name).Type
            Next idxFld
            Exit Function
        End If
    Next idx

End Function


'---------------------------------------------------------------------------------------
' Function  : GetTableMergeStrategy
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Returns how a merge can bring this table's data in line with its source
'           : file, setting strReason to a translated explanation when it cannot.
'           :
'           : Binary, complex, and calculated columns rule out any strategy. Binary and
'           : complex values cannot be compared with a SQL operator, and calculated
'           : columns are maintained by the engine and cannot be assigned. Those tables
'           : still export and still import on a full build.
'           :
'           : With a merge key the rows are reconciled individually. Without one, a source
'           : row cannot be matched to a table row, so the only option is to replace every
'           : row -- acceptable precisely because there is no key, and therefore no
'           : identity or AutoNumber value that anything could be holding a reference to.
'           : It is still refused when a relationship points at the table, since the
'           : delete would fail and the whole table would roll back.
'---------------------------------------------------------------------------------------
'
Public Function GetTableMergeStrategy(tdf As DAO.TableDef, _
    ByRef strReason As String) As eTableMergeStrategy

    Dim fld As DAO.Field
    Dim strDependent As String

    strReason = vbNullString

    For Each fld In tdf.Fields
        If IsBinaryTableFieldType(fld.Type) Then
            strReason = T("binary field '{0}'", var0:=fld.Name)
            Exit Function
        End If
    Next fld

    ' Covers complex (multi-value/attachment) and calculated fields, both of which
    ' depend on the embedded XML schema to round-trip.
    If TableRequiresXmlSchema(tdf) Then
        strReason = T("complex or calculated fields")
        Exit Function
    End If

    If GetTableMergeKey(tdf).Count > 0 Then
        GetTableMergeStrategy = etmsReconcile
        Exit Function
    End If

    strDependent = GetFirstDependentTable(tdf.Name)
    If Len(strDependent) > 0 Then
        strReason = T("no primary key or unique required index, and '{0}' references it", _
            var0:=strDependent)
        Exit Function
    End If

    GetTableMergeStrategy = etmsReload

End Function


'---------------------------------------------------------------------------------------
' Function  : GetFirstDependentTable
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Return the name of a table that references this one through a relationship,
'           : or an empty string when nothing does. Only the first is needed; it exists to
'           : name a blocking table in a message.
'---------------------------------------------------------------------------------------
'
Public Function GetFirstDependentTable(strTable As String) As String

    Dim rel As DAO.Relation

    For Each rel In SharedDb.Relations
        ' Table is the referenced (one) side, ForeignTable the referencing (many) side.
        If StrComp(rel.Table, strTable, vbTextCompare) = 0 Then
            If StrComp(rel.ForeignTable, strTable, vbTextCompare) <> 0 Then
                GetFirstDependentTable = rel.ForeignTable
                Exit Function
            End If
        End If
    Next rel

End Function


'---------------------------------------------------------------------------------------
' Function  : IsBinaryTableFieldType
' Author    : Adam Waller
' Date      : 7/27/2026
' Purpose   : Returns true for DAO field types that cannot be represented in table data
'           : export text/XML formats.
'---------------------------------------------------------------------------------------
'
Public Function IsBinaryTableFieldType(intType As Integer) As Boolean
    Select Case intType
        Case dbLongBinary, dbVarBinary, dbAttachment: IsBinaryTableFieldType = True
        Case Else: IsBinaryTableFieldType = False
    End Select
End Function


'---------------------------------------------------------------------------------------
' Function  : TableRequiresXmlSchema
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Returns true when the embedded XML schema has to survive sanitization for
'           : the table to import correctly. Calculated fields carry od:expression, and
'           : complex or OLE object fields carry an od:jetType that ImportXML needs.
'           : Exporting through a query drops those annotations, so these tables must be
'           : exported with acExportTable instead of a sorted query.
'---------------------------------------------------------------------------------------
'
Public Function TableRequiresXmlSchema(tdf As DAO.TableDef) As Boolean

    Dim fld As DAO.Field
    Dim strExpression As String

    For Each fld In tdf.Fields

        Select Case fld.Type
            Case dbLongBinary, dbAttachment, _
                dbComplexByte, dbComplexInteger, dbComplexLong, dbComplexSingle, _
                dbComplexDouble, dbComplexGUID, dbComplexDecimal, dbComplexText
                TableRequiresXmlSchema = True
                Exit Function
        End Select

        ' Only calculated fields expose an Expression property, so reading it from an
        ' ordinary field raises an error that we use as the negative result.
        strExpression = vbNullString
        LogUnhandledErrors
        On Error Resume Next
        strExpression = Nz(fld.Properties("Expression"), vbNullString)
        If Err Then Err.Clear
        On Error GoTo 0
        If Len(strExpression) > 0 Then
            TableRequiresXmlSchema = True
            Exit Function
        End If

    Next fld

End Function


'---------------------------------------------------------------------------------------
' Procedure : PrepareTableDataSortExport
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Prepare for a table-data export operation by removing any temporary sort
'           : query left behind by an interrupted run. Pairs with
'           : ReleaseTableDataSortExport, which drops the query this operation creates.
'---------------------------------------------------------------------------------------
'
Public Sub PrepareTableDataSortExport()
    SweepLeftoverTableDataSortQueries
End Sub


'---------------------------------------------------------------------------------------
' Function  : AssignTableDataSortQuery
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Point the operation-scoped temporary sort query at strSql. Creates the
'           : query on first use and reassigns .SQL on later tables, which avoids paying
'           : a QueryDefs.Refresh over the whole collection once per table.
'           :
'           : Returns an empty string when strSql has no ORDER BY, or when the query
'           : could not be created or repointed. Returning empty on failure is essential:
'           : the query would otherwise still hold the previous table's SQL, and the
'           : caller would export that table's rows into this table's source file.
'           :
'           : Note that the engine defers table-name resolution, so assigning SQL that
'           : names a missing table succeeds here and fails later in ExportXML, where the
'           : caller already falls back to the table export.
'---------------------------------------------------------------------------------------
'
Public Function AssignTableDataSortQuery(strSql As String) As String

    Dim dbs As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim strName As String

    If InStr(1, strSql, " ORDER BY ", vbTextCompare) = 0 Then Exit Function

    Set dbs = SharedDb

    Perf.OperationStart "Assign Temp Sort Query"

    LogUnhandledErrors
    On Error Resume Next

    ' Repoint the query we already own.
    If Len(m_strTableDataSortQueryName) > 0 Then
        Set qdf = dbs.QueryDefs(m_strTableDataSortQueryName)
        If Err Then
            ' Gone from under us (interrupted run, external cleanup). Recreate below.
            Err.Clear
            m_strTableDataSortQueryName = vbNullString
        Else
            qdf.SQL = strSql
            If Err Then GoTo Failed
        End If
    End If

    ' Create on first use, or to replace one that disappeared.
    If Len(m_strTableDataSortQueryName) = 0 Then
        strName = GetUnusedTableDataSortQueryName
        dbs.CreateQueryDef strName, strSql
        If Err Then GoTo Failed
        ' Needed once per operation so ExportXML can resolve the new name.
        dbs.QueryDefs.Refresh
        If Err Then GoTo Failed
        m_strTableDataSortQueryName = strName
    End If

    On Error GoTo 0
    AssignTableDataSortQuery = m_strTableDataSortQueryName
    Perf.OperationEnd
    Exit Function

Failed:
    ' Drop the query rather than leave it pointed at a previous table's SQL.
    Err.Clear
    If Len(m_strTableDataSortQueryName) > 0 Then
        dbs.QueryDefs.Delete m_strTableDataSortQueryName
        If Err Then Err.Clear
        m_strTableDataSortQueryName = vbNullString
    End If
    On Error GoTo 0
    Perf.OperationEnd

End Function


'---------------------------------------------------------------------------------------
' Procedure : ReleaseTableDataSortExport
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Drop the operation-scoped temporary sort query. Safe to call when no query
'           : is active, so it can sit unconditionally in export cleanup blocks.
'---------------------------------------------------------------------------------------
'
Public Sub ReleaseTableDataSortExport()

    Dim dbs As DAO.Database

    If Len(m_strTableDataSortQueryName) = 0 Then Exit Sub

    Perf.OperationStart "Drop Temp Sort Query"
    Set dbs = SharedDb

    LogUnhandledErrors
    On Error Resume Next
    dbs.QueryDefs.Delete m_strTableDataSortQueryName
    If Err Then Err.Clear
    On Error GoTo 0

    m_strTableDataSortQueryName = vbNullString
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SweepLeftoverTableDataSortQueries
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Remove any temporary sort queries left behind by an interrupted export.
'---------------------------------------------------------------------------------------
'
Private Sub SweepLeftoverTableDataSortQueries()

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim strName As String

    Set dbs = SharedDb

    LogUnhandledErrors
    On Error Resume Next
    Set rst = dbs.OpenRecordset( _
        "SELECT Name FROM MSysObjects WHERE Name LIKE '" & TABLE_DATA_SORT_QUERY_PREFIX & "*' AND Type = 5", _
        dbOpenSnapshot, dbReadOnly)
    If Not rst Is Nothing Then
        Do While Not rst.EOF
            strName = Nz(rst!Name, vbNullString)
            If Len(strName) > 0 Then
                dbs.QueryDefs.Delete strName
                If Err Then Err.Clear
            End If
            rst.MoveNext
        Loop
        rst.Close
        Set rst = Nothing
    End If
    If Err Then Err.Clear
    On Error GoTo 0

End Sub


'---------------------------------------------------------------------------------------
' Function  : GetUnusedTableDataSortQueryName
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Return a query name not already used by any object in the database.
'---------------------------------------------------------------------------------------
'
Private Function GetUnusedTableDataSortQueryName() As String

    Dim strName As String
    Dim lngSuffix As Long

    strName = TABLE_DATA_SORT_QUERY_PREFIX
    Do While DCount("*", "MSysObjects", "Name=""" & strName & """") > 0
        lngSuffix = lngSuffix + 1
        strName = TABLE_DATA_SORT_QUERY_PREFIX & CStr(lngSuffix)
    Loop

    GetUnusedTableDataSortQueryName = strName

End Function


'---------------------------------------------------------------------------------------
' Function  : CreateTableDataStagingTable
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Create an empty local table with the same columns as strTable, to load a
'           : source file into before reconciling it against the live table. Returns the
'           : staging table name, or an empty string when it could not be created.
'           :
'           : SELECT INTO is used rather than DDL so every column keeps the source
'           : table's exact type and size. It also demotes an AutoNumber key to a plain
'           : Long, which is what we want: the staging copy has to hold the key values
'           : read from the source file.
'           :
'           : A unique index on the merge key is added for two reasons. The engine
'           : rejects an UPDATE across a join unless the joined side is provably unique,
'           : and the index is what keeps the reconcile joins from scanning. A table with
'           : no merge key gets no index, since it is reloaded wholesale rather than
'           : joined, and a unique index could reject rows the live table accepts.
'---------------------------------------------------------------------------------------
'
Public Function CreateTableDataStagingTable(strTable As String) As String

    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim dKey As Dictionary
    Dim varKey As Variant
    Dim cFields As clsConcat
    Dim cKeys As clsConcat
    Dim strName As String

    ' An interrupted merge can leave one behind, and table defs get exported.
    Perf.OperationStart "Sweep Staging Tables"
    SweepLeftoverTableDataStagingTables
    Perf.OperationEnd

    Set dbs = SharedDb
    Set tdf = dbs.TableDefs(strTable)
    Set dKey = GetTableMergeKey(tdf)

    Set cFields = New clsConcat
    For Each fld In tdf.Fields
        cFields.Add "[", fld.Name, "], "
    Next fld
    cFields.Remove 2

    Set cKeys = New clsConcat
    For Each varKey In dKey.Keys
        cKeys.Add "[", CStr(varKey), "], "
    Next varKey
    If cKeys.Length > 0 Then cKeys.Remove 2

    strName = GetUnusedTableDataStagingName

    Perf.OperationStart "Create Staging Table"
    LogUnhandledErrors
    On Error Resume Next
    dbs.Execute "SELECT " & cFields.GetStr & " INTO [" & strName & "] FROM [" & strTable & _
        "] WHERE (1 = 0)", dbFailOnError
    If Err Then GoTo Failed
    If cKeys.Length > 0 Then
        dbs.Execute "CREATE UNIQUE INDEX [idx_" & strName & "] ON [" & strName & "] (" & _
            cKeys.GetStr & ")", dbFailOnError
        If Err Then GoTo Failed
    End If
    ' The collection was enumerated before this table existed.
    dbs.TableDefs.Refresh
    If Err Then GoTo Failed
    On Error GoTo 0

    CreateTableDataStagingTable = strName
    Perf.OperationEnd
    Exit Function

Failed:
    Err.Clear
    On Error GoTo 0
    Perf.OperationEnd
    ' Never return a half-built staging table for the caller to load rows into.
    DropTableDataStagingTable strName

End Function


'---------------------------------------------------------------------------------------
' Procedure : DropTableDataStagingTable
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Remove a staging table. Safe to call with an empty name or for a table that
'           : was never created, so it can sit unconditionally in a cleanup block.
'---------------------------------------------------------------------------------------
'
Public Sub DropTableDataStagingTable(strName As String)

    Dim dbs As DAO.Database

    If Len(strName) = 0 Then Exit Sub
    Set dbs = SharedDb

    LogUnhandledErrors
    On Error Resume Next
    dbs.Execute "DROP TABLE [" & strName & "]", dbFailOnError
    If Err Then Err.Clear
    dbs.TableDefs.Refresh
    If Err Then Err.Clear
    On Error GoTo 0

End Sub


'---------------------------------------------------------------------------------------
' Procedure : SweepLeftoverTableDataStagingTables
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Remove staging tables left behind by an interrupted merge. Only ever called
'           : before a new staging table is created, so this can never drop one that is
'           : currently in use.
'---------------------------------------------------------------------------------------
'
Private Sub SweepLeftoverTableDataStagingTables()

    Dim dbs As DAO.Database
    Dim rst As DAO.Recordset
    Dim colNames As Collection
    Dim varName As Variant

    Set dbs = SharedDb
    Set colNames = New Collection

    LogUnhandledErrors
    On Error Resume Next
    Set rst = dbs.OpenRecordset( _
        "SELECT Name FROM MSysObjects WHERE Name LIKE '" & TABLE_DATA_STAGING_PREFIX & "*' AND Type = 1", _
        dbOpenSnapshot, dbReadOnly)
    If Not rst Is Nothing Then
        Do While Not rst.EOF
            colNames.Add Nz(rst!Name, vbNullString)
            rst.MoveNext
        Loop
        rst.Close
        Set rst = Nothing
    End If
    If Err Then Err.Clear
    On Error GoTo 0

    For Each varName In colNames
        DropTableDataStagingTable CStr(varName)
    Next varName

End Sub


'---------------------------------------------------------------------------------------
' Function  : GetUnusedTableDataStagingName
' Author    : Adam Waller
' Date      : 7/28/2026
' Purpose   : Return a table name not already used by any object in the database.
'---------------------------------------------------------------------------------------
'
Private Function GetUnusedTableDataStagingName() As String

    Dim strName As String
    Dim lngSuffix As Long

    strName = TABLE_DATA_STAGING_PREFIX
    Do While DCount("*", "MSysObjects", "Name=""" & strName & """") > 0
        lngSuffix = lngSuffix + 1
        strName = TABLE_DATA_STAGING_PREFIX & CStr(lngSuffix)
    Loop

    GetUnusedTableDataStagingName = strName

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetCachedTableType
' Author    : Adam Waller
' Date      : 7/13/2026
' Purpose   : Return the cached MSysObjects.Type for a table name when a batch cache is
'           : active. Returns 0 when no cache is active or the name is not cached.
'---------------------------------------------------------------------------------------
'
Public Function GetCachedTableType(strName As String) As Long

    If m_dTableTypeCache Is Nothing Then Exit Function
    If Not m_dTableTypeCache.Exists(strName) Then Exit Function
    GetCachedTableType = m_dTableTypeCache(strName)

End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildTableTypeCache
' Author    : Adam Waller
' Date      : 7/13/2026
' Purpose   : Batch-load every local/linked table Type from MSysObjects into a module
'           : cache in a single recordset pass, replacing hundreds of per-table Type
'           : lookups during a table scan. Always rebuilds fresh so a cache from a
'           : prior operation can never go stale. System and temporary tables are
'           : skipped to match clsDbTableDef.GetAllFromDB enumeration.
'---------------------------------------------------------------------------------------
'
Public Sub BuildTableTypeCache()

    Dim dbs As Database
    Dim rst As DAO.Recordset
    Dim strName As String
    Dim dSysTables As Dictionary

    Perf.OperationStart "Build Table Type Cache"

    Set m_dTableTypeCache = New Dictionary
    m_dTableTypeCache.CompareMode = TextCompare

    If DebugMode(True) Then On Error GoTo Err_Handler Else On Error Resume Next

    Set dbs = SharedDb
    Set dSysTables = GetSystemTableNames
    Set rst = dbs.OpenRecordset( _
        "SELECT Name, Type FROM MSysObjects WHERE Type IN (1,4,6)", _
        dbOpenSnapshot, dbReadOnly)
    If rst Is Nothing Then GoTo Err_Handler

    Do While Not rst.EOF
        strName = Nz(rst!Name, vbNullString)
        If Len(strName) > 0 Then
            If Not (dSysTables.Exists(strName) Or strName Like "~*") Then
                m_dTableTypeCache(strName) = Nz(rst!Type, 1)
            End If
        End If
        rst.MoveNext
    Loop
    rst.Close

    ' Clear any benign leftover error so the failure handler only fires on real faults.
    Err.Clear
    On Error GoTo 0
    GoTo CleanUp

Err_Handler:
    On Error Resume Next
    ' On any failure, drop the (possibly partial) cache so IsLinkedTable falls back to the
    ' reliable per-table Type read rather than trusting incomplete data.
    CatchAny eelWarning, "Error building table Type cache", _
        ModuleName & ".BuildTableTypeCache", True, True
    Set m_dTableTypeCache = Nothing

CleanUp:
    Set rst = Nothing
    Perf.OperationEnd

End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearTableTypeCache
' Author    : Adam Waller
' Date      : 7/13/2026
' Purpose   : Release the batched table Type cache after a scan completes. Safe to call
'           : when no cache is active.
'---------------------------------------------------------------------------------------
'
Public Sub ClearTableTypeCache()
    Set m_dTableTypeCache = Nothing
End Sub


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
