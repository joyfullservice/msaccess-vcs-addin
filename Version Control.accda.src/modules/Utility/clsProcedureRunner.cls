VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProcedureRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Author    : Adam Waller, Josef Pötzl
' Date      : 4/23/2020
' Purpose   : Encapsulates the execution of procedures within the current Microsoft Access project.
'           : The class provides a dedicated and extensible abstraction for procedure execution,
'           : keeping this functionality separate from database-related concerns.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
'@Folder("Utility")

Private Const ModuleName As String = "clsProcedureRunner"

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
    Dim arrParams() As Variant

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
