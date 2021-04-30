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

' Logging and options classes
Private m_Perf As clsPerformance
Private m_Log As clsLog
Private m_Options As clsOptions
Private m_VCSIndex As clsVCSIndex

' Keep a persistent reference to file system object after initializing version control.
' This way we don't have to recreate this object dozens of times while using VCS.
Private m_FSO As Scripting.FileSystemObject


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
    If m_Options Is Nothing Then Set m_Options = LoadOptions
    Set Options = m_Options
End Property
Public Property Set Options(cNewOptions As clsOptions)
    Set m_Options = cNewOptions
End Property


'---------------------------------------------------------------------------------------
' Procedure : Perf
' Author    : Adam Waller
' Date      : 11/3/2020
' Purpose   : Wrapper for performance logging class
'---------------------------------------------------------------------------------------
'
Public Function Perf() As clsPerformance
    If m_Perf Is Nothing Then Set m_Perf = New clsPerformance
    Set Perf = m_Perf
End Function


'---------------------------------------------------------------------------------------
' Procedure : Log
' Author    : Adam Waller
' Date      : 4/28/2020
' Purpose   : Wrapper for log file class
'---------------------------------------------------------------------------------------
'
Public Function Log() As clsLog
    If m_Log Is Nothing Then Set m_Log = New clsLog
    Set Log = m_Log
End Function


'---------------------------------------------------------------------------------------
' Procedure : FSO
' Author    : Adam Waller
' Date      : 1/18/2019
' Purpose   : Wrapper for file system object. A property allows us to clear the object
'           : reference when we have completed an export or import operation.
'---------------------------------------------------------------------------------------
'
Public Property Get FSO() As Scripting.FileSystemObject
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next
    If m_FSO Is Nothing Then Set m_FSO = New Scripting.FileSystemObject
    Set FSO = m_FSO
    CatchAny eelCritical, "Unable to create Scripting.FileSystemObject", ModuleName & ".FSO"
End Property
Public Property Set FSO(ByVal RHS As Scripting.FileSystemObject)
    Set m_FSO = RHS
End Property


'---------------------------------------------------------------------------------------
' Procedure : VSCIndex
' Author    : Adam Waller
' Date      : 12/1/2020
' Purpose   : Reference to the VCS Index class (saved state from vcs-index.json)
'---------------------------------------------------------------------------------------
'
Public Property Get VCSIndex() As clsVCSIndex
    If m_VCSIndex Is Nothing Then
        Set m_VCSIndex = New clsVCSIndex
        m_VCSIndex.LoadFromFile
    End If
    Set VCSIndex = m_VCSIndex
End Property
Public Property Set VCSIndex(cIndex As clsVCSIndex)
    Set m_VCSIndex = cIndex
End Property


'---------------------------------------------------------------------------------------
' Procedure : DebugMode
' Author    : Adam Waller
' Date      : 3/9/2021
' Purpose   : Wrapper for use in error handling.
'---------------------------------------------------------------------------------------
'
Public Function DebugMode(blnTrapUnhandledErrors As Boolean) As Boolean
    
    Dim blnBreak As Boolean
    
    ' Don't reference the property this till we have loaded the options.
    If Not m_Options Is Nothing Then blnBreak = m_Options.BreakOnError
    
    ' Check for any unhandled errors
    If (Err.Number <> 0) And blnTrapUnhandledErrors Then
    
        ' Check current BreakOnError mode
        If blnBreak Then
            ' Stop the code here so we can investigate the source of the error.
            Debug.Print "Error " & Err.Number & ": " & Err.Description
            Stop
            '===========================================================================
            '   NOTE: IF THE CODE STOPS HERE, PLEASE READ BEFORE CONTINUING
            '===========================================================================
            '   An unhandled error was (probably) found just before an `On Error ...`
            '   statement. Since any existing errors are cleared when the On Error
            '   statement is executed, this is your chance to identify the source of the
            '   unhandled error.
            '
            '   Note that the error will typically be from the THIRD item in the call
            '   stack, if the On Error statement is at the beginning of the calling
            '   procedure. Use CTL+L to view the call stack. For example:
            '
            '   (1) MSAccessVCS.modObjects.DebugMode    <--- This function
            '   (2) MSAccessVCS.clsLog.Flush            <--- Calling function
            '   (3) MSAccessVCS.clsLog.Add              <--- Likely origin of error
            '
            '   You can use standard VBA debugging techniques to inspect variables and
            '   step through code to pinpoint the source and cause of the error.
            '   For additional information, please see the add-in wiki on GitHub at:
            '   https://github.com/joyfullservice/msaccess-vcs-integration/wiki
            '===========================================================================
        Else
            ' Log otherwise unhandled error
            If Not m_Log Is Nothing Then
                ' We don't know the procedure that it originated from, but we should at least
                ' log that the error occurred. A review of the log file may help identify the source.
                Log.Error eelError, "Unhandled error found before `On Error` directive", "Unknown"
            End If
        End If
    
    End If
    
    ' Return debug mode
    DebugMode = blnBreak
    
End Function