Attribute VB_Name = "modTestWebRunner"
'---------------------------------------------------------------------------------------
' Module    : modTestWebRunner
' Author    : Adam Waller
' Date      : 8/21/2026
' Purpose   : Unit tests for web test-runner bridge helpers (timeout detection, JSON
'           : shape guards). No EdgeBrowserControl required.
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests")
'@Tag("unit")


Private Const ENGLISH_TIMEOUT As String = _
    "RetrieveJavascriptValue timed out. Please verify the JavaScript expression supplied is valid."
Private Const GERMAN_TIMEOUT As String = _
    "Timeout bei RetrieveJavascriptValue. Bitte überprüfen Sie, ob der angegebene JavaScript-Ausdruck gültig ist."
Private Const ITALIAN_TIMEOUT As String = _
    "Timeout di RetrieveJavascriptValue. Verificare che l'espressione JavaScript fornita sia valida."


Public Sub TestJsRetrieveTimeout_English()
    TestAssert IsJsRetrieveTimeout(ENGLISH_TIMEOUT), "English timeout sentinel"
End Sub


Public Sub TestJsRetrieveTimeout_German()
    TestAssert IsJsRetrieveTimeout(GERMAN_TIMEOUT), "German timeout sentinel"
End Sub


Public Sub TestJsRetrieveTimeout_Italian()
    TestAssert IsJsRetrieveTimeout(ITALIAN_TIMEOUT), "Italian timeout sentinel"
End Sub


Public Sub TestJsRetrieveTimeout_ValidJsonNotTimeout()
    TestAssert Not IsJsRetrieveTimeout("[]"), "empty array is not a timeout"
    TestAssert Not IsJsRetrieveTimeout("[{""fn"":""RunAll""}]"), "outbox JSON is not a timeout"
End Sub


Public Sub TestJsRetrieveTimeout_TypeofResultsNotTimeout()
    TestAssert Not IsJsRetrieveTimeout("object"), "typeof object result is not a timeout"
    TestAssert Not IsJsRetrieveTimeout("true"), "boolean string result is not a timeout"
End Sub


Public Sub TestJsonRootStart_AcceptsObjectAndArray()
    TestAssert IsJsonRootStart("{}"), "empty object"
    TestAssert IsJsonRootStart("[]"), "empty array"
    TestAssert IsJsonRootStart("[{""fn"":""Cancel""}]"), "outbox array"
    TestAssert IsJsonRootStart("{""o"":[],""d"":[]}"), "diag wrapper object"
End Sub


Public Sub TestJsonRootStart_RejectsNonJson()
    TestAssert Not IsJsonRootStart(ENGLISH_TIMEOUT), "English timeout is not JSON"
    TestAssert Not IsJsonRootStart(GERMAN_TIMEOUT), "German timeout is not JSON"
    TestAssert Not IsJsonRootStart("object"), "typeof result is not JSON"
    TestAssert Not IsJsonRootStart("true"), "boolean string is not JSON"
    TestAssert Not IsJsonRootStart(vbNullString), "empty string is not JSON"
End Sub
