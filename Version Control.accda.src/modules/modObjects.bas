Attribute VB_Name = "modObjects"
'---------------------------------------------------------------------------------------
' Module    : modObjects
' Author    : Adam Waller
' Date      : 12/4/2020
' Purpose   : Wrapper functions for classes and other objects available globally.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit

Private Const ModuleName = "modObjects"

' Use a private type to manage instances of object classes
Private Type udtObjects
    Perf As clsPerformance
    Log As clsLog
    Options As clsOptions
    VCSIndex As clsVCSIndex
    Worker As clsWorker
    Git As clsGitIntegration
    Translation As clsTranslation

    ' Keep a persistent reference to file system object after initializing version control.
    ' This way we don't have to recreate this object dozens of times while using VCS.
    FSO As Scripting.FileSystemObject
End Type
Private this As udtObjects


'---------------------------------------------------------------------------------------
' Procedure : ReleaseObjects
' Author    : Adam Waller
' Date      : 3/28/2022
' Purpose   : Release references to objects for a clean exit.
'---------------------------------------------------------------------------------------
'
Public Sub ReleaseObjects()
    Set this.Perf = Nothing
    Set this.Log = Nothing
    Set this.Options = Nothing
    Set this.VCSIndex = Nothing
    Set this.Worker = Nothing
    Set this.Git = Nothing
    Set this.FSO = Nothing

    Dim udtEmpty As udtObjects
    ' Reassign "this" to blank, clearing any saved data.
    LSet this = udtEmpty
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LoadOptions
' Author    : Adam Waller
' Date      : 4/15/2020
' Purpose   : Loads the current options from defaults and this project.
'---------------------------------------------------------------------------------------
'
Public Function LoadOptions() As clsOptions
    Dim Options As clsOptions
    Set Options = New clsOptions
    Options.LoadProjectOptions
    Set LoadOptions = Options
End Function


'---------------------------------------------------------------------------------------
' Procedure : Options
' Author    : Adam Waller
' Date      : 5/2/2020
' Purpose   : A global property to access options from anywhere in code.
'           : (Avoiding a global state is better OO programming, but this approach keeps
'           :  the coding simpler when you don't have to tie everything back to the
'           :  primary object.)
'           : To clear the current set of options, simply set the property to nothing.
'---------------------------------------------------------------------------------------
'
Public Property Get Options() As clsOptions
    If this.Options Is Nothing Then Set this.Options = LoadOptions
    Set Options = this.Options
End Property
Public Property Set Options(cNewOptions As clsOptions)
    Set this.Options = cNewOptions
End Property


'---------------------------------------------------------------------------------------
' Procedure : OptionsLoaded
' Author    : Adam Waller
' Date      : 5/13/2023
' Purpose   : Return true if the options object has been loaded. (It is loaded on first
'           : access, but in some cases we might want to avoid loading the options if
'           : they are not already loaded.)
'---------------------------------------------------------------------------------------
'
Public Property Get OptionsLoaded() As Boolean
    OptionsLoaded = (Not this.Options Is Nothing)
End Property


'---------------------------------------------------------------------------------------
' Procedure : Perf
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Wrapper for performance logging class
'---------------------------------------------------------------------------------------
'
Public Function Perf() As clsPerformance
    If this.Perf Is Nothing Then Set this.Perf = New clsPerformance
    Set Perf = this.Perf
End Function


'---------------------------------------------------------------------------------------
' Procedure : Translation
' Author    : Adam Waller
' Date      : 5/19/2021
' Purpose   : Expose translation class
'---------------------------------------------------------------------------------------
'
Public Function Translation() As clsTranslation
    If this.Translation Is Nothing Then Set this.Translation = New clsTranslation
    Set Translation = this.Translation
End Function


'---------------------------------------------------------------------------------------
' Procedure : T
' Author    : Adam Waller
' Date      : 3/19/2024
' Purpose   : Wrapper function to translate to current language
'---------------------------------------------------------------------------------------
'
Public Function T(strText As String, Optional strComments As String, Optional strContext As String, _
    Optional var0, Optional var1, Optional var2, Optional var3, Optional var4, _
    Optional var5, Optional var6, Optional var7, Optional var8, Optional var9)
    T = Translation.T(strText, strComments, strContext, _
        var0, var1, var2, var3, var4, var5, var6, var7, var8, var9)
End Function


'---------------------------------------------------------------------------------------
' Procedure : Log
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Wrapper for log file class
'---------------------------------------------------------------------------------------
'
Public Function Log(Optional blnCreateInstance As Boolean = True) As clsLog
    If this.Log Is Nothing Then If blnCreateInstance Then Set this.Log = New clsLog
    Set Log = this.Log
End Function


'---------------------------------------------------------------------------------------
' Procedure : FSO
' Author    : Adam Waller, hecon5
' Date      : 1/18/2019, 10/24/2024
' Purpose   : Wrapper for file system object. A property allows us to clear the object
'           : reference when we have completed an export or import operation.
'---------------------------------------------------------------------------------------
'
Public Property Get FSO() As Scripting.FileSystemObject

    Const FunctionName As String = ModuleName & ".FSO"
    Static RetryCount As Long

    LogUnhandledErrors FunctionName
    On Error Resume Next

Retry:
    If this.FSO Is Nothing Then Set this.FSO = New Scripting.FileSystemObject
    Set FSO = this.FSO
    If CatchAny(eelError, "Retry FSO Check", FunctionName, False, True) And RetryCount < 2 Then
        ' Some machines in some environments may fail to generate the FileSystemObject the first time
        ' 99% of the time, the second attempt will work. This may be due to a race condition in the OS.
        RetryCount = RetryCount + 1
        GoTo Retry
    End If
    CatchAny eelCritical, "Unable to create Scripting.FileSystemObject", FunctionName

End Property


'---------------------------------------------------------------------------------------
' Procedure : VSCIndex
' Author    : Adam Waller
' Date      : 12/1/2020
' Purpose   : Reference to the VCS Index class (saved state from vcs-index.json)
'---------------------------------------------------------------------------------------
'
Public Property Get VCSIndex() As clsVCSIndex
    If this.VCSIndex Is Nothing Then
        Set this.VCSIndex = New clsVCSIndex
        this.VCSIndex.LoadFromFile
    End If
    Set VCSIndex = this.VCSIndex
End Property
Public Property Set VCSIndex(cIndex As clsVCSIndex)
    Set this.VCSIndex = cIndex
End Property


'---------------------------------------------------------------------------------------
' Procedure : Worker
' Author    : Adam Waller
' Date      : 3/2/2023
' Purpose   : Expose worker class
'---------------------------------------------------------------------------------------
'
Public Property Get Worker() As clsWorker
    If this.Worker Is Nothing Then Set this.Worker = New clsWorker
    Set Worker = this.Worker
End Property


'---------------------------------------------------------------------------------------
' Procedure : Diff
' Author    : Adam Waller
' Date      : 2/23/2022
' Purpose   : Wrapper for class to view diff between string and file content
'---------------------------------------------------------------------------------------
'
Public Property Get Diff() As clsViewDiff
    Static cDiff As clsViewDiff
    If cDiff Is Nothing Then Set cDiff = New clsViewDiff
    Set Diff = cDiff
End Property


'---------------------------------------------------------------------------------------
' Procedure : Git
' Author    : Adam Waller
' Date      : 3/10/2023
' Purpose   : Return Git integration class
'---------------------------------------------------------------------------------------
'
Public Property Get Git() As clsGitIntegration
    If this.Git Is Nothing Then Set this.Git = New clsGitIntegration
    Set Git = this.Git
End Property
