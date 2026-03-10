Attribute VB_Name = "modErrorHandling"
'---------------------------------------------------------------------------------------
' Module    : modErrorHandling
' Author    : Adam Waller
' Date      : 5/13/2023
' Purpose   : General error handling functions.
' Layer     : Infrastructure
' Depends on: modObjects (Log singleton only — single documented coupling point)
' Note      : BreakOnError is cached locally via ConfigureErrorHandling() to avoid
'           : a circular dependency with modObjects/Options. The Log.Error coupling
'           : remains because VBA lacks function pointers for object methods.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Private Module
Option Explicit
'@Folder("Infrastructure")

Private Const ModuleName As String = "modErrorHandling"

Private Type udtThis
    blnInError As Boolean       ' Monitor error state
    blnBreakOnError As Boolean  ' Cached from Options; set via ConfigureErrorHandling
End Type
Private this As udtThis


'---------------------------------------------------------------------------------------
' Procedure : DebugMode
' Author    : Adam Waller
' Date      : 4/14/2023
' Purpose   : Wrapper for use in error handling.
'---------------------------------------------------------------------------------------
'
Public Function DebugMode(blnTrapUnhandledErrors As Boolean) As Boolean

    ' Log any unhandled errors
    If blnTrapUnhandledErrors Then LogUnhandledErrors

    ' Use cached value set by ConfigureErrorHandling (defaults to False
    ' until modObjects loads options and pushes the setting here).
    DebugMode = this.blnBreakOnError

End Function


'---------------------------------------------------------------------------------------
' Procedure : ConfigureErrorHandling
' Author    : Adam Waller
' Date      : 3/10/2026
' Purpose   : Called by modObjects after options load to push the BreakOnError setting
'           : into this module. This avoids a circular dependency: modErrorHandling no
'           : longer needs to reach into Options (modObjects) to read BreakOnError.
'---------------------------------------------------------------------------------------
'
Public Sub ConfigureErrorHandling(blnBreakOnError As Boolean)
    this.blnBreakOnError = blnBreakOnError
End Sub


'---------------------------------------------------------------------------------------
' Procedure : LogUnhandledErrors
' Author    : Adam Waller
' Date      : 4/14/2023
' Purpose   : Log any unhandled error condition, also breaking code execution if that
'           : option is currently set. (Run this before any ON ERROR directive which
'           : will siently reset any current VBA error condition.)
'
' Example   : See Sub `CatchTest` for example use.
'
'---------------------------------------------------------------------------------------
'
Public Sub LogUnhandledErrors(Optional ByRef CallingFunction As String = vbNullString)

    ' Check for any unhandled errors
    If (Err.Number <> 0) And Not this.blnInError Then

        this.blnInError = True ' Set flag so we don't create a loop while logging the error

        ' Check cached BreakOnError flag (pushed by modObjects.ConfigureErrorHandling)
        If this.blnBreakOnError Then
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
            '   (1) MSAccessVCS.modErrorHandling.DebugMode  <--- This function
            '   (2) MSAccessVCS.clsLog.Flush                <--- Calling function
            '   (3) MSAccessVCS.clsLog.Add                  <--- Likely origin of error
            '
            '   You can use standard VBA debugging techniques to inspect variables and
            '   step through code to pinpoint the source and cause of the error.
            '   For additional information, please see the add-in wiki on GitHub at:
            '   https://github.com/joyfullservice/msaccess-vcs-addin/wiki
            '===========================================================================
        Else
            ' Log otherwise unhandled error
            ' We don't know the procedure that it originated from, but we should at least
            ' log that the error occurred. A review of the log file may help identify the source.
            Log.Error eelError, "Unhandled error, likely before `On Error` directive", CallingFunction & ".Unknown.LogUnhandledErrors"
        End If
        this.blnInError = False
    End If

End Sub


'---------------------------------------------------------------------------------------
' Procedure : Catch
' Author    : Adam Waller
' Date      : 11/23/2020
' Purpose   : Returns true if the last error matches any of the passed error numbers,
'           : and clears the error object.
'---------------------------------------------------------------------------------------
'
Public Function Catch(ParamArray lngErrorNumbers()) As Boolean
    Dim intCnt As Integer
    For intCnt = LBound(lngErrorNumbers) To UBound(lngErrorNumbers)
        If lngErrorNumbers(intCnt) = Err.Number Then
            Err.Clear
            Catch = True
            Exit For
        End If
    Next intCnt
End Function


'---------------------------------------------------------------------------------------
' Procedure : CatchAny
' Author    : Adam Waller
' Date      : 12/3/2020
' Purpose   : Generic error handler with logging.
'---------------------------------------------------------------------------------------
'
Public Function CatchAny(eLevel As eErrorLevel, strDescription As String, Optional strSource As String, _
    Optional blnLogError As Boolean = True, Optional blnClearError As Boolean = True, _
    Optional blnIncludeErrorWithDescription As Boolean = False) As Boolean
    If Err Then
        If blnLogError Then
            this.blnInError = True
            Log.Error eLevel, strDescription, strSource, blnIncludeErrorWithDescription
            this.blnInError = False
        End If
        If blnClearError Then Err.Clear
        CatchAny = True
    End If
End Function
