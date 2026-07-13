VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDbCommandBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller
' Date      : 4/23/2020
' Purpose   : This class extends the IDbComponent class to perform the specific
'           : operations required by this particular object type.
'           : (I.e. The specific way you export or import this component.)
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Components")

Private m_CommandBar As CommandBar
Private m_TemplateCommandBar As CommandBar
Private m_Items(True To False) As Dictionary
Private m_FileList As Dictionary

Private Type udtThis
    BarPosition As MsoBarPosition
    BarProperties As Collection
    CtlReadProperties As Collection
    CtlWriteProperties As Collection
    dImages As Dictionary       ' Dictionary of image objects used with this CommandBar
End Type
Private this As udtThis

' This requires us to use all the public methods and properties of the implemented class
' which keeps all the component classes consistent in how they are used in the export
' and import process. The implemented functions should be kept private as they are called
' from the implementing class, not this class.
Implements IDbComponent


'---------------------------------------------------------------------------------------
' Procedure : Export
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Export the individual database component (table, form, query, etc...)
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Export(Optional strAlternatePath As String)
    Dim strContent As String
    strContent = GetSource
    WriteFile strContent, Nz2(strAlternatePath, IDbComponent_SourceFile)
    SaveImages
    VCSIndex.Update Me, IIf(strAlternatePath = vbNullString, eatExport, eatAltExport), GetStringHash(strContent, True)
End Sub


'---------------------------------------------------------------------------------------
' Procedure : SaveImages
' Author    : Adam Waller
' Date      : 1/12/2024
' Purpose   : Saves the images for the menu bar
'---------------------------------------------------------------------------------------
'
Private Sub SaveImages()

    Dim varPic As Variant

    ' Make sure we actually have images to save
    If this.dImages Is Nothing Then Exit Sub
    If this.dImages.Count = 0 Then Exit Sub

    ' Verify the first path, to make sure the folder exists
    VerifyPath CStr(this.dImages.Keys(0))

    Perf.OperationStart "Save CommandBar Images"
    For Each varPic In this.dImages.Keys
        ' Save the item image to a file
        stdole.SavePicture this.dImages(varPic).Picture, varPic & "_Picture.bmp"
        ' Also save image mask
        stdole.SavePicture this.dImages(varPic).Mask, varPic & "_Mask.bmp"
    Next varPic
    Perf.OperationEnd

    ' After saving the images, we can clear the references
    ' to the image objects
    Set this.dImages = Nothing

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Import
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Import the individual database component from a file.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Import(strFile As String)

    Dim dBar As Dictionary
    Dim strName As String

    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' Only import files with the correct extension.
    If Not strFile Like "*.json" Then Exit Sub
    Set dBar = ReadJsonFile(strFile)

    Perf.OperationStart "Building CommandBar"
    ' Delete any existing bar with this name.
    strName = GetObjectNameFromFileName(strFile)
    Set m_CommandBar = GetExisting(strName)
    If Not m_CommandBar Is Nothing Then
        If m_CommandBar.BuiltIn Then
            Log.Error eelError, "You cannot replace the existing built-in command bar: " & m_CommandBar.Name, ModuleName(Me) & ".Import"
            Exit Sub
        Else
            ' Delete the command bar
            m_CommandBar.Delete
            Set m_CommandBar = Nothing
        End If
    End If

    ' Now, create a new command bar, and add controls to it.
    Dim IsMenu As Boolean
    Dim Position As MsoBarPosition

    Position = -1 'set to invalid value to simulate the "missing" value which cannot be directly set in VBA...
    IsMenu = (dBar("Items").Item("Type") = msoBarTypeMenuBar)
    If CatchAny(eelNoError, vbNullString) Then
        IsMenu = False
    End If
    Position = dBar("Items").Item("Position")
    If CatchAny(eelNoError, vbNullString) Then
        Position = -1
    End If

    If Position = -1 Then
        Set m_CommandBar = CommandBars.Add(strName, , IsMenu, False)
    Else
        Set m_CommandBar = CommandBars.Add(strName, Position, IsMenu, False)
    End If
    BuildControls dBar("Items"), m_CommandBar
    Perf.OperationEnd

    ' Log any errors
    CatchAny eelError, "Importing CommandBar " & strName, ModuleName(Me) & ".Import"

    ' Save to index
    VCSIndex.Update Me, eatImport, GetDictionaryHash(GetDictionary)

End Sub


'---------------------------------------------------------------------------------------
' Procedure : BuildControls
' Author    : Adam Waller
' Date      : 1/13/2024
' Purpose   : Recursive function to build the controls on the popup menu
'---------------------------------------------------------------------------------------
'
Private Sub BuildControls(dParent As Dictionary, objItem As Object)

    Dim varProp As Variant
    Dim dControl As Dictionary
    Dim picItem As IPictureDisp
    Dim strBasePath As String
    Dim strPath As String

    LogUnhandledErrors
    On Error Resume Next

    If TypeOf objItem Is CommandBar Then
        ' Set command bar properties
        For Each varProp In this.BarProperties
            If varProp <> "Type" Then
                If dParent.Exists(varProp) Then
                    CallByName objItem, varProp, VbLet, dParent(varProp)
                    If Catch(438, 440) Then
                        ' No point logging those errors.
                    ElseIf CatchAny(eelNoError, vbNullString) Then
                        Debug.Print "Could not set " & varProp & " on " & objItem.Name
                    End If
                End If
            End If
        Next
        ' Now add any nested controls (built-in handling shared with popups via AddChildControl)
        If dParent.Exists("Controls") Then
            For Each dControl In dParent("Controls")
                AddChildControl dControl, objItem
            Next dControl
        End If
    ElseIf TypeOf objItem Is CommandBarControl Then
        ' Set control properties
        For Each varProp In this.CtlWriteProperties
            If dParent.Exists(varProp) Then
                CallByName objItem, varProp, VbLet, dParent(varProp)
                If Catch(438, 440) Then
                    ' No point logging those errors
                ElseIf CatchAny(eelNoError, vbNullString) Then
                    Debug.Print "Could not set " & varProp & " on " & objItem.Caption
                End If
            End If
        Next varProp
        If TypeOf objItem Is CommandBarPopup And dParent.Exists("Controls") Then
            ' Add nested controls (built-in controls handled the same as at the top level)
            For Each dControl In dParent("Controls")
                AddChildControl dControl, objItem
            Next dControl
        ElseIf TypeOf objItem Is CommandBarComboBox And dParent.Exists("List") Then
            Dim varItem As Variant
            Dim cbcItem As CommandBarComboBox
            Set objItem = cbcItem
            For Each varItem In dParent("List")
                cbcItem.AddItem varItem
            Next
        Else
            ' Check for picture
            If dParent.Exists("ImagePath") Then
                strBasePath = GetPathFromRelative(dParent("ImagePath"), IDbComponent_BaseFolder)

                ' Build out full path to main image (Picture)
                strPath = strBasePath & "_Picture.bmp"
                If FSO.FileExists(strPath) Then
                    Set picItem = stdole.LoadPicture(strPath)
                    objItem.Picture = picItem
                End If
                ' Build path to image mask (transparency)
                strPath = strBasePath & "_Mask.bmp"
                If FSO.FileExists(strPath) Then
                    Set picItem = stdole.LoadPicture(strPath)
                    objItem.Mask = picItem
                End If
            End If
        End If
    Else
        ' Unsupported object type
        Stop
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddChildControl
' Author    : Adam Waller
' Date      : 6/22/2026
' Purpose   : Add a single child control to a parent (CommandBar or CommandBarPopup),
'           : applying the correct strategy for built-in, Access "custom built-in",
'           : and fully custom controls. Shared by both nesting levels so built-in
'           : controls round-trip identically whether at the top level or in a submenu.
'---------------------------------------------------------------------------------------
'
Private Sub AddChildControl(dControl As Dictionary, objParent As Object)

    Dim objBuiltIn As CommandBarControl
    Dim objCustom As CommandBarControl
    Dim objTemplateCtl As CommandBarControl
    Dim objCopiedCtl As CommandBarControl
    Dim objDestBar As CommandBar

    LogUnhandledErrors
    On Error Resume Next

    If dControl("BuiltIn") Then
        Set objBuiltIn = objParent.Controls.Add(dControl("Type"), dControl("Id"))
        If Catch(438, 440, -2147467259) Then
            ' Expected for non-addable built-ins; recreate from the template command bar.
            Set objTemplateCtl = Nothing
            Set objCopiedCtl = Nothing
            If m_TemplateCommandBar Is Nothing Then
                WarnTemplateCommandBarNotLoaded
            Else
                Set objTemplateCtl = m_TemplateCommandBar.FindControl(Id:=dControl("Id"))
                CatchAny eelNoError, vbNullString
                If objTemplateCtl Is Nothing Then
                    Log.Error eelError, "Built-in control Id " & dControl("Id") & " could not be recreated (not addable and not in the template command bar).", ModuleName(Me) & ".AddChildControl"
                Else
                    ' CommandBarControl.Copy targets a CommandBar; for a popup parent copy into its submenu.
                    If TypeOf objParent Is CommandBar Then
                        Set objDestBar = objParent
                    Else
                        Set objDestBar = objParent.CommandBar
                    End If
                    Set objCopiedCtl = objTemplateCtl.Copy(objDestBar)
                    If Not CatchAny(eelError, "Unable to copy built-in control Id " & dControl("Id") & " from the template", ModuleName(Me) & ".AddChildControl", blnIncludeErrorWithDescription:=True) Then
                        If Not objCopiedCtl Is Nothing Then BuildControls dControl, objCopiedCtl
                    End If
                End If
            End If
        ElseIf CatchAny(eelError, "Unable to add built-in control Id " & dControl("Id"), ModuleName(Me) & ".AddChildControl", blnIncludeErrorWithDescription:=True) Then
            ' Unexpected error adding a built-in that should be addable.
        Else
            objBuiltIn.Visible = dControl("Visible")
            If dControl.Exists("BeginGroup") Then objBuiltIn.BeginGroup = dControl("BeginGroup")
        End If
    Else
        ' Fully custom controls and non-addable built-in replicas (BuiltIn=false).
        Set objCustom = objParent.Controls.Add(dControl("Type"))
        If CatchAny(eelError, "Unable to add custom control to command bar", ModuleName(Me) & ".AddChildControl", blnIncludeErrorWithDescription:=True) Then
            ' Failed to add custom control.
        Else
            BuildControls dControl, objCustom
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : WarnTemplateCommandBarNotLoaded
' Author    : Adam Waller
' Date      : 6/30/2026
' Purpose   : Log once per class instance when the add-in template bar is absent from
'           : Application.CommandBars (expected after rebuilding the add-in in the same
'           : session). Avoids per-control errors and raw Err leakage from .Add failures.
'---------------------------------------------------------------------------------------
'
Private Sub WarnTemplateCommandBarNotLoaded()

    Static blnWarned As Boolean

    If blnWarned Then Exit Sub
    blnWarned = True
    Log.Error eelWarning, T("The command bar template is not loaded in this Access session (expected after rebuilding the add-in). Restart Access before building databases with non-addable built-in menu controls."), ModuleName(Me) & ".AddChildControl"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetControlFaceId
' Author    : Adam Waller
' Date      : 6/23/2026
' Purpose   : Read FaceId, which only exists on button controls. Combo boxes, edit
'           : boxes, and popups don't expose it, so we skip the access we know would
'           : fail rather than triggering and swallowing a 438. The error path is
'           : reserved for genuine unknowns, flagged as warnings so they get addressed.
'---------------------------------------------------------------------------------------
'
Private Function GetControlFaceId(ctl As CommandBarControl) As Long
    If Not (TypeOf ctl Is CommandBarButton) Then Exit Function
    LogUnhandledErrors
    On Error Resume Next
    GetControlFaceId = ctl.FaceId
    CatchAny eelWarning, "Unexpected error reading FaceId for button '" & ctl.Caption & _
        "' (Id " & ctl.Id & ")", ModuleName(Me) & ".GetControlFaceId"
End Function


'---------------------------------------------------------------------------------------
' Procedure : SkipControlImage
' Author    : Adam Waller
' Date      : 6/22/2026
' Purpose   : Returns true when a control's image should NOT be exported to a file.
'           : Under export format 5.0.0+, only genuinely custom (pasted) faces are saved
'           : as images. A non-zero FaceId is a built-in icon that is reconstructed from
'           : the FaceId on import, so exporting its rendered bitmap only creates churn.
'---------------------------------------------------------------------------------------
'
Private Function SkipControlImage(ctl As CommandBarControl) As Boolean
    If Options.ExportFormatVersion < EFV_5_0_0 Then Exit Function
    SkipControlImage = (GetControlFaceId(ctl) <> 0)
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsBuiltInControlCustomized
' Author    : Adam Waller
' Date      : 6/25/2026
' Purpose   : Return true when a built-in control carries customizations that the minimal
'           : built-in export (Type + Id only) would lose: a custom action (OnAction /
'           : Parameter) or a custom pasted face (FaceId = 0 with a picture present). Such
'           : controls must be exported as full replicas so their action parameters and
'           : pictures survive the round-trip (issue #584). A pristine built-in control
'           : has an empty OnAction - its behavior is wired to its Id - so a non-empty
'           : OnAction/Parameter reliably indicates a user customization.
'---------------------------------------------------------------------------------------
'
Private Function IsBuiltInControlCustomized(ctl As CommandBarControl) As Boolean

    Dim picItem As IPictureDisp

    LogUnhandledErrors
    On Error Resume Next

    ' Custom action assigned to the built-in control.
    If Len(ctl.OnAction) > 0 Then IsBuiltInControlCustomized = True
    If Len(ctl.Parameter) > 0 Then IsBuiltInControlCustomized = True
    CatchAny eelNoError, vbNullString
    If IsBuiltInControlCustomized Then Exit Function

    ' Custom pasted face: a button with no FaceId but a picture present.
    If TypeOf ctl Is CommandBarButton Then
        If GetControlFaceId(ctl) = 0 Then
            Set picItem = ctl.Picture
            If Catch(438) Then
                ' No picture available; not a custom face.
            ElseIf CatchAny(eelNoError, vbNullString) Then
                ' Unexpected error; treat as not customized.
            ElseIf Not picItem Is Nothing Then
                IsBuiltInControlCustomized = True
            End If
        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : IsTemplateControl
' Author    : Adam Waller
' Date      : 6/25/2026
' Purpose   : Return true when the built-in control Id is present on the add-in's template
'           : command bar. Non-addable controls in the template are exported as full-property
'           : built-ins and recreated by template copy on import. When the template is not
'           : loaded (e.g. consecutive add-in builds), returns False and export degrades to
'           : a visual replica.
'---------------------------------------------------------------------------------------
'
Private Function IsTemplateControl(lngId As Long) As Boolean
    If m_TemplateCommandBar Is Nothing Then Exit Function
    LogUnhandledErrors
    On Error Resume Next
    IsTemplateControl = Not (m_TemplateCommandBar.FindControl(Id:=lngId) Is Nothing)
    CatchAny eelNoError, vbNullString
End Function


'---------------------------------------------------------------------------------------
' Procedure : AddReferenceName
' Author    : Adam Waller
' Date      : 6/23/2026
' Purpose   : Add a read-only, human/agent reference name. Ignored on import. Names
'           : come from ControlIdToName (never the live Caption) so export stays
'           : deterministic. Two distinct keys make the meaning explicit:
'           :  - Built-in controls get "ControlIdName" (keyed off the control Id),
'           :    which reflects the control's actual identity/purpose.
'           :  - Custom controls get "FaceIdName" (keyed off the FaceId), which only
'           :    describes the borrowed icon's appearance, not the control's purpose.
'---------------------------------------------------------------------------------------
'
Private Sub AddReferenceName(dParent As Dictionary, objItem As CommandBarControl)

    Dim strName As String
    Dim lngFaceId As Long

    If Options.ExportFormatVersion < EFV_5_0_0 Then Exit Sub

    If objItem.BuiltIn Then
        If dParent.Exists("ControlIdName") Then Exit Sub
        strName = ControlIdToName(objItem.Id)
        If Len(strName) = 0 Then
            Log.Add "  Unknown built-in control Id " & objItem.Id & " (caption: '" & objItem.Caption & "')", False
        Else
            dParent.Add "ControlIdName", strName
        End If
    Else
        lngFaceId = GetControlFaceId(objItem)
        If lngFaceId <> 0 Then
            If dParent.Exists("FaceIdName") Then Exit Sub
            strName = ControlIdToName(lngFaceId)
            If Len(strName) > 0 Then dParent.Add "FaceIdName", strName
        End If
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddReplicaReferenceName
' Author    : Adam Waller
' Date      : 6/24/2026
' Purpose   : Add a FaceIdName breadcrumb for built-in controls exported as custom
'           : replicas (EFV_5_0_0+). Uses FaceId when present, otherwise the control Id.
'---------------------------------------------------------------------------------------
'
Private Sub AddReplicaReferenceName(dParent As Dictionary, objItem As CommandBarControl)

    Dim lngFaceId As Long
    Dim strName As String

    If Options.ExportFormatVersion < EFV_5_0_0 Then Exit Sub
    If dParent.Exists("FaceIdName") Then Exit Sub

    lngFaceId = GetControlFaceId(objItem)
    If lngFaceId = 0 Then lngFaceId = objItem.Id
    strName = ControlIdToName(lngFaceId)
    If Len(strName) > 0 Then dParent.Add "FaceIdName", strName

End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetExisting
' Author    : Adam Waller
' Date      : 1/13/2024
' Purpose   : Returns any existing command bar by that name
'---------------------------------------------------------------------------------------
'
Private Function GetExisting(strName As String) As CommandBar
    LogUnhandledErrors
    On Error Resume Next
    Set GetExisting = CommandBars(strName)
    If Err Then Err.Clear
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetSource
' Author    : Adam Waller
' Date      : 2/14/2022
' Purpose   : Return the full content that will be saved to the source file.
'---------------------------------------------------------------------------------------
'
Private Function GetSource() As String
    GetSource = BuildJsonFile(TypeName(Me), GetDictionary, "CommandBar")
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetDictionary
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a dictionary object of project properties.
'---------------------------------------------------------------------------------------
'
Private Function GetDictionary() As Dictionary
    Dim dParent As Dictionary
    Set dParent = New Dictionary
    Set this.dImages = New Dictionary
    Set GetDictionary = BuildElementDictionary(dParent, m_CommandBar)
End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildElementDictionary
' Author    : Adam Waller
' Date      : 1/12/2024
' Purpose   : A recursive function to build out the dictionary elements representing
'           : a CommandBar popup menu.
'           : objItem may represent a CommandBar, or a CommandBarControl.
'---------------------------------------------------------------------------------------
'
Private Function BuildElementDictionary(dParent As Dictionary, objItem As Object) As Dictionary

    Dim varProp As Variant
    Dim varValue As Variant
    Dim colItems As Collection
    Dim picItem As IPictureDisp
    Dim strPath As String
    Dim cbcItem As CommandBarComboBox
    Dim colList As Collection
    Dim lngItem As Long
    Dim blnExportAsReplica As Boolean

    LogUnhandledErrors
    On Error Resume Next

    If TypeOf objItem Is CommandBar Then
        ' Get command bar position
        this.BarPosition = objItem.Position
        ' Add command bar properties
        For Each varProp In this.BarProperties
            varValue = CallByName(objItem, varProp, VbGet)
            If Not CatchAny(eelNoError, vbNullString) Then
                dParent.Add varProp, varValue
            End If
        Next varProp
        ' Now add any nested controls
        Set colItems = New Collection
        dParent.Add "Controls", BuildControlCollection(objItem.Controls)
    ElseIf TypeOf objItem Is CommandBarControl Then
        blnExportAsReplica = False
        ' Built-in classification (EFV_5_0_0+): manual override -> replica; addable+clean ->
        ' minimal block; addable+customized -> replica (#584); non-addable+in-template ->
        ' full-property built-in (template copy on import); non-addable+not-in-template ->
        ' visual replica. Addability uses a cached runtime probe keyed on (Type, Id).
        If objItem.BuiltIn Then
            If Options.ExportFormatVersion >= EFV_5_0_0 Then
                If IsNonAddableControl(objItem.Id) Then
                    blnExportAsReplica = True
                ElseIf IsBuiltInControlAddable(objItem.Type, objItem.Id) Then
                    If IsBuiltInControlCustomized(objItem) Then
                        blnExportAsReplica = True
                    Else
                        AddReferenceName dParent, objItem
                        dParent.Add "BeginGroup", objItem.BeginGroup
                        dParent.Add "BuiltIn", True
                        dParent.Add "Type", objItem.Type
                        dParent.Add "Visible", objItem.Visible
                        dParent.Add "Id", objItem.Id
                        Set BuildElementDictionary = dParent
                        Exit Function
                    End If
                ElseIf IsTemplateControl(objItem.Id) Then
                    ' Non-addable but present on the template bar -> full built-in export.
                Else
                    blnExportAsReplica = True
                End If
            End If
        End If
        If blnExportAsReplica Then
            AddReplicaReferenceName dParent, objItem
        ElseIf Options.ExportFormatVersion >= EFV_5_0_0 Then
            AddReferenceName dParent, objItem
        End If
        ' Add control properties
        For Each varProp In this.CtlReadProperties
            varValue = Empty ' Clear the value from previous run
            varValue = CallByName(objItem, varProp, VbGet)
            ' No point in logging those 2 errors.
            If Catch(438, 440) Then
            ' Check for any other errors which should be logged just in case.
            ElseIf Not CatchAny(eelNoError, vbNullString) Then
                dParent.Add varProp, varValue
            End If
        Next varProp
        If blnExportAsReplica Then
            If dParent.Exists("BuiltIn") Then
                dParent("BuiltIn") = False
            Else
                dParent.Add "BuiltIn", False
            End If
            If dParent.Exists("Id") Then dParent.Remove "Id"
        ElseIf objItem.BuiltIn Then
            dParent.Add "Id", objItem.Id
        End If
        If TypeOf objItem Is CommandBarComboBox Then
            If objItem.Type = msoControlDropdown Or objItem.Type = msoControlComboBox Then
            ' Build List
            Set cbcItem = objItem

            Set colList = New Collection
            For lngItem = 1 To cbcItem.ListCount
                colList.Add cbcItem.List(lngItem), CStr(lngItem)
            Next
            dParent.Add "List", colList
            End If
        End If
        If TypeOf objItem Is CommandBarPopup Then
            ' Addable built-in popups have their native submenus regenerated on import via
            ' .Add(Type, Id), so serializing their children would create duplicates. But a
            ' non-addable built-in popup is exported as a replica and rebuilt from source,
            ' so its children must be serialized. Custom popups always serialize children.
            If Options.ExportFormatVersion >= EFV_5_0_0 And objItem.BuiltIn And Not blnExportAsReplica Then
                ' Skip children for addable built-in popups (regenerated on import)
            Else
                dParent.Add "Controls", BuildControlCollection(objItem.Controls)
            End If
        Else
            ' Check for picture
            Set picItem = objItem.Picture
            ' Not all command bar controls have picture
            If Catch(438) Then
                ' Don't log this error and skip
            ElseIf CatchAny(eelNoError, vbNullString) Then
                ' Log any other error and skip
            Else
                ' Proceed with picture processing (only for genuinely custom faces;
                ' a non-zero FaceId is a built-in icon rebuilt from FaceId on import)
                If Not picItem Is Nothing And Not SkipControlImage(objItem) Then
                    strPath = BuildPath2(IDbComponent_BaseFolder, GetImagePath(objItem))
                    ' Check path for possible duplicates (menu items with the same name)
                    If this.dImages.Exists(strPath) Then
                        ' Add a number to make the path unique.
                        strPath = strPath & "_" & this.dImages.Count
                    End If
                    ' Save reference to image objects to use when
                    ' exporting images to files.
                    this.dImages.Add strPath, objItem
                    ' Save path to images in element dictionary
                    dParent.Add "ImagePath", GetRelativePath(strPath, IDbComponent_BaseFolder)
                End If
            End If
        End If
    Else
        ' Unsupported object type
        Stop
    End If

    ' Remove relative position unless this is a floating command bar
    If Options.ExportFormatVersion >= EFV_5_0_0 Then
        If TypeOf objItem Is CommandBar And this.BarPosition = msoBarFloating Then
            ' Preserve Top and Left position for the floating command bar
        Else
            ' Remove cached coordinate positions that will change frequently
            If dParent.Exists("Left") Then dParent.Remove "Left"
            If dParent.Exists("Top") Then dParent.Remove "Top"
            ' RowIndex (docking row) drifts the same way for non-floating bars
            If dParent.Exists("RowIndex") Then dParent.Remove "RowIndex"
        End If
    End If

    ' Return dictionary
    Set BuildElementDictionary = dParent

End Function


'---------------------------------------------------------------------------------------
' Procedure : BuildControlCollection
' Author    : Adam Waller
' Date      : 1/12/2024
' Purpose   : Return a collection of the command bar controls.
'---------------------------------------------------------------------------------------
'
Private Function BuildControlCollection(ctls As CommandBarControls) As Collection
    Dim ctl As CommandBarControl
    Dim dItem As Dictionary
    Set BuildControlCollection = New Collection
    For Each ctl In ctls
        Set dItem = New Dictionary
        BuildControlCollection.Add BuildElementDictionary(dItem, ctl)
    Next ctl
End Function


'---------------------------------------------------------------------------------------
' Procedure : GetImagePath
' Author    : Adam Waller
' Date      : 1/12/2024
' Purpose   : Build a base path for the image files. (Image and mask)
'---------------------------------------------------------------------------------------
'
Private Function GetImagePath(ctl As CommandBarControl) As String

    Dim ctlTest As Object
    Dim strName As String
    Dim strSegment As String

    On Error GoTo 0
    Set ctlTest = ctl
    ' Walk up the parent objects to build a full path to the control.
    ' (We are using a name path since the control IDs and indexes may change)
    Do
        If TypeOf ctlTest Is CommandBar Then
            ' Don't include command bar name, since we are using a subfolder
            ' to store images for each commmand bar.
            strSegment = vbNullString
        ElseIf TypeOf ctlTest Is CommandBarPopup Then
            ' We already have the name for this item
            strSegment = vbNullString
        Else
            ' Control item. (Use caption or ID)
            strSegment = Nz2(MultiReplace(ctlTest.Caption, _
                "&", vbNullString, _
                "...", vbNullString) _
                , ctlTest.Id)
        End If
        If Len(strSegment) Then
            If Len(strName) Then
                strName = strSegment & "_" & strName
            Else
                strName = strSegment
            End If
        End If
        If ctlTest.Parent Is Application Then
            ' This is the top level
            Exit Do
        Else
            ' Move up to parent object
            Set ctlTest = ctlTest.Parent
        End If
    Loop

    ' Convert name to a filesafe name
    GetImagePath = BuildPath2(GetSafeFileName(m_CommandBar.Name) & "_Images", GetSafeFileName(strName))

End Function


'---------------------------------------------------------------------------------------
' Procedure : Merge
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Merge the source file into the existing database, updating or replacing
'           : any existing object.
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_Merge(strFile As String)
    ' Import if file exists
    If FSO.FileExists(strFile) Then
        IDbComponent_Import strFile
    Else
        VCSIndex.Remove Me, strFile
    End If
End Sub


'---------------------------------------------------------------------------------------
' Procedure : IDbComponent_MoveSource
' Author    : Adam Waller
' Date      : 9/10/2022
' Purpose   : Move the component's source file(s) from one folder to another
'---------------------------------------------------------------------------------------
'
Private Sub IDbComponent_MoveSource(strFromFolder As String, strToFolder As String)
    MoveFileIfExists strFromFolder & FSO.GetFileName(IDbComponent_SourceFile), strToFolder
    MoveFolderIfExists strFromFolder & FSO.GetBaseName(IDbComponent_SourceFile), strToFolder
End Sub


'---------------------------------------------------------------------------------------
' Procedure : GetAllFromDB
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a collection of class objects represented by this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetAllFromDB(Optional blnModifiedOnly As Boolean = False) As Dictionary

    Dim cCmdBar As IDbComponent
    Dim bar As CommandBar
    Dim blnNeedAll As Boolean
    Dim sngStart As Single
    Dim lngScanCount As Long

    ' Build collection if not already cached
    If m_Items(blnModifiedOnly) Is Nothing Then
        ' Loop through command bars, looking for any custom ones.

        ' TODO: Figure out a way to determine whether a command bar belongs to the current
        ' database. The command bars are associated with the application, not with the database
        ' so the application could have additional command bars loaded via add-in or other processes
        ' that are not actually a part of the database itself. The data is there in the internal
        ' table but not human readable.
        blnNeedAll = (m_Items(False) Is Nothing)
        If blnNeedAll Then Set m_Items(False) = New Dictionary
        If blnModifiedOnly Then Set m_Items(True) = New Dictionary
        If blnModifiedOnly Then
            lngScanCount = Application.CommandBars.Count
            sngStart = Timer
        End If
        For Each bar In Application.CommandBars
            If bar.BuiltIn Then
                ' Skip
            ElseIf bar.Name = strTemplateCommandBarName Then
                ' Skip our template command bar
            Else
                Set cCmdBar = New clsDbCommandBar
                Set cCmdBar.DbObject = bar
                If blnNeedAll Then m_Items(False).Add cCmdBar.SourceFile, cCmdBar
                If blnModifiedOnly Then
                    If cCmdBar.IsModified Then m_Items(True).Add cCmdBar.SourceFile, cCmdBar
                    Log.IncrementObjectScanProgress sngStart, lngScanCount, "CommandBars"
                End If
            End If
        Next bar

    End If

    ' Return cached collection
    Set IDbComponent_GetAllFromDB = m_Items(blnModifiedOnly)

End Function


'---------------------------------------------------------------------------------------
' Procedure : GetFileList
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a list of file names to import for this component type.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_GetFileList() As Dictionary
    If m_FileList Is Nothing Then Set m_FileList = GetFilePathsInFolder(IDbComponent_BaseFolder, "*.json")
    Set IDbComponent_GetFileList = m_FileList
End Function


'---------------------------------------------------------------------------------------
' Procedure : IsModified
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Returns true if the object in the database has been modified since
'           : the last export of the object.
'---------------------------------------------------------------------------------------
'
Public Function IDbComponent_IsModified() As Boolean
    IDbComponent_IsModified = VCSIndex.Item(Me).FileHash <> GetStringHash(GetSource, True)
End Function


'---------------------------------------------------------------------------------------
' Procedure : DateModified
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : The date/time the object was modified. (If possible to retrieve)
'           : If the modified date cannot be determined (such as application
'           : properties) then this function will return 0.
'---------------------------------------------------------------------------------------
'
Private Function IDbComponent_DateModified() As Date
    IDbComponent_DateModified = 0
End Function


'---------------------------------------------------------------------------------------
' Procedure : Category
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a category name for this type. (I.e. forms, queries, macros)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Category() As String
    IDbComponent_Category = "CommandBars"
End Property


'---------------------------------------------------------------------------------------
' Procedure : BaseFolder
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return the base folder for import/export of this component.
'---------------------------------------------------------------------------------------
Private Property Get IDbComponent_BaseFolder() As String
    IDbComponent_BaseFolder = Options.GetExportFolder & "menus" & PathSep
End Property


'---------------------------------------------------------------------------------------
' Procedure : FileExtensions
' Author    : Adam Waller
' Date      : 12/1/2023
' Purpose   : A collection of the file extensions used in source files for this
'           : component type.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_FileExtensions() As Collection
    Set IDbComponent_FileExtensions = New Collection
    IDbComponent_FileExtensions.Add "json"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Name
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a name to reference the object for use in logs and screen output.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Name() As String
    If m_CommandBar Is Nothing Then Exit Property
    IDbComponent_Name = m_CommandBar.Name
End Property


'---------------------------------------------------------------------------------------
' Procedure : SourceFile
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return the full path of the source file for the current object.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SourceFile() As String
    If m_CommandBar Is Nothing Then Exit Property
    IDbComponent_SourceFile = IDbComponent_BaseFolder & GetSafeFileName(m_CommandBar.Name) & ".json"
End Property


'---------------------------------------------------------------------------------------
' Procedure : Count
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Return a count of how many items are in this category.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_Count(Optional blnModifiedOnly As Boolean = False) As Long
    IDbComponent_Count = IDbComponent_GetAllFromDB(blnModifiedOnly).Count
End Property


'---------------------------------------------------------------------------------------
' Procedure : QuickCount
' Author    : Adam Waller
' Date      : 6/14/2022
' Purpose   : Return a cached, non-iterative approximate count of database objects
'           : for use with progress indicators when scanning for changes. Single file
'           : objects like database properties can simply return 1.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_QuickCount() As Long
    ' Probably 0 in most databases. We will have to iterate through the CommandBars
    ' to check the .BuiltIn flag to get an actual count.
    IDbComponent_QuickCount = 0
End Property


'---------------------------------------------------------------------------------------
' Procedure : ComponentType
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : The type of component represented by this class.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_ComponentType() As eDatabaseComponentType
    IDbComponent_ComponentType = edbCommandBar
End Property


'---------------------------------------------------------------------------------------
' Procedure : DbObject
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : This represents the database object we are dealing with.
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_DbObject() As Object
    Set IDbComponent_DbObject = m_CommandBar
End Property
Private Property Set IDbComponent_DbObject(ByVal RHS As Object)
    Set m_CommandBar = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : SingleFile
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Returns true if the export of all items is done as a single file instead
'           : of individual files for each component. (I.e. properties, references)
'---------------------------------------------------------------------------------------
'
Private Property Get IDbComponent_SingleFile() As Boolean
    IDbComponent_SingleFile = False
End Property


'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : Adam Waller
' Date      : 5/17/2021
' Purpose   : Helps us know whether we have already counted the objects.
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()

    ' The template may not be available when building the add-in from source
    ' multiple times in one session. AfterBuild's ImportCommandBarsTemplate
    ' deletes all instances from Application.CommandBars before reimporting,
    ' removing the add-in's library copy permanently for the session.
    ' The template is restored on the next Access restart.
    LogUnhandledErrors
    On Error Resume Next

    ' We assume the add-in file already has the template command bar included. If it doesn't, then
    ' we will have problem importing Access' custom built-in buttons.
    Set m_TemplateCommandBar = Application.CommandBars(strTemplateCommandBarName)
    If CatchAny(eelNoError, vbNullString) Then
        Log.Add "  Template command bar not available (expected during consecutive add-in builds)", False
    End If

    ' Set up property name lists

    ' CommandBar Properties
    Set this.BarProperties = New Collection
    With this.BarProperties
        .Add "AdaptiveMenu"
        .Add "Context"
        .Add "Enabled"
        .Add "Left"
        .Add "NameLocal"
        .Add "Position"
        .Add "Protection"
        .Add "RowIndex"
        .Add "Top"
        .Add "Type"
    End With

    ' CommmandBarControl Properties
    ' This should include all properties from different types of CommandBarControl
    ' even if they are not common among all types.

    ' Build a list for read-write properties to be used in building controls.
    Set this.CtlWriteProperties = New Collection
    With this.CtlWriteProperties
        .Add "BeginGroup"
        .Add "BuiltInFace"
        .Add "Caption"
        .Add "DescriptionText"
        .Add "DropDownLines"
        .Add "DropDownWidth"
        .Add "Enabled"
        .Add "FaceId"
        .Add "HelpContextId"
        .Add "HelpFile"
        .Add "HyperlinkType"
        .Add "ListHeaderCount"
        .Add "ListIndex"
        .Add "OLEMenuGroup"
        .Add "OLEUsage"
        .Add "OnAction"
        .Add "Parameter"
        .Add "ShortcutText"
        .Add "Style"
        .Add "Tag"
        .Add "Text"
        .Add "TooltipText"
        .Add "Visible"
    End With

    ' Build dictionary to be used for writing out to text files to include additional data about the
    ' controls so any changes can be detected via source code control
    Set this.CtlReadProperties = New Collection
    With this.CtlReadProperties
        .Add "BeginGroup"
        .Add "BuiltIn"
        .Add "BuiltInFace"
        .Add "Caption"
        .Add "DescriptionText"
        .Add "DropDownLines"
        .Add "DropDownWidth"
        .Add "Enabled"
        .Add "FaceId"
        .Add "HelpContextId"
        .Add "HelpFile"
        .Add "HyperlinkType"
        .Add "Left"
        .Add "ListCount"
        .Add "ListHeaderCount"
        .Add "ListIndex"
        .Add "OLEMenuGroup"
        .Add "OLEUsage"
        .Add "OnAction"
        .Add "Parameter"
        .Add "ShortcutText"
        .Add "Style"
        .Add "Tag"
        .Add "Text"
        .Add "TooltipText"
        .Add "Top"
        .Add "Type"
        .Add "Visible"
    End With

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Class_Terminate
' Author    : Adam Waller
' Date      : 1/13/2024
' Purpose   : Clear any object references
'---------------------------------------------------------------------------------------
'
Private Sub Class_Terminate()
    ' Clear the images dictionary, just in case it still has references
    ' to the commandbar image objects.
    Set this.dImages = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure : Parent
' Author    : Adam Waller
' Date      : 4/24/2020
' Purpose   : Return a reference to this class as an IDbComponent. This allows you
'           : to reference the public methods of the parent class without needing
'           : to create a new class object.
'---------------------------------------------------------------------------------------
'
Public Property Get Parent() As IDbComponent
    Set Parent = Me
End Property
