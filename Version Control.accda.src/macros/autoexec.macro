Version =196611
ColumnsShown =0
Begin
    Action ="SetLocalVar"
    Argument ="__*L0_"
    Argument ="[CurrentProject].[IsTrusted]"
End
Begin
    Action ="SetLocalVar"
    Argument ="__*L0C_"
    Argument ="[LocalVars]![__*L0_]"
End
Begin
    Condition ="[LocalVars]![__*L0C_]"
    Action ="RunCode"
    Argument ="AutoRun()"
End
Begin
    Action ="SetLocalVar"
    Argument ="__*L0C_"
    Argument ="(Not [LocalVars]![__*L0_])"
End
Begin
    Condition ="[LocalVars]![__*L0C_]"
    Action ="MsgBox"
    Argument ="This add-In file is not currently trusted. Please click \"Enable Content\" to tr"
        "ust the add-in and run the startup code normally."
    Argument ="-1"
    Argument ="2"
    Argument ="Add-in is not trusted"
End
Begin
    Action ="SetLocalVar"
    Argument ="__*L0C_"
    Argument ="False"
End
Begin
    Comment ="_AXL:<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"no\"?>\015\012<UserI"
        "nterfaceMacro MinimumClientDesignVersion=\"14.0.0000.0000\" xmlns=\"http://schem"
        "as.microsoft.com/office/accessservices/2009/11/application\" xmlns:a=\"http://sc"
        "hemas.microsoft.com/office/acc"
End
Begin
    Comment ="_AXL:essservices/2009/11/forms\"><Statements><ConditionalBlock><If><Condition>[C"
        "urrentProject].[IsTrusted]</Condition><Statements><Action Name=\"RunCode\"><Argu"
        "ment Name=\"FunctionName\">AutoRun()</Argument></Action></Statements></If><Else>"
        "<Statements><Actio"
End
Begin
    Comment ="_AXL:n Name=\"MessageBox\"><Argument Name=\"Message\">This add-In file is not cu"
        "rrently trusted. Please click \"Enable Content\" to trust the add-in and run the"
        " startup code normally.</Argument><Argument Name=\"Type\">Warning?</Argument><Ar"
        "gument Name=\"Title\">A"
End
Begin
    Comment ="_AXL:dd-in is not trusted</Argument></Action></Statements></Else></ConditionalBl"
        "ock></Statements></UserInterfaceMacro>"
End
