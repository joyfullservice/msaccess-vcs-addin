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
    Argument ="The add-In is not trusted currently. Please click \"Enable Contents\" to trust t"
        "he add-in and thus execute the startup code normally."
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
        "as.microsoft.com/office/accessservices/2009/11/application\"><Statements><Condit"
        "ionalBlock><If><Condition>[Cu"
End
Begin
    Comment ="_AXL:rrentProject].[IsTrusted]</Condition><Statements><Action Name=\"RunCode\"><"
        "Argument Name=\"FunctionName\">AutoRun()</Argument></Action></Statements></If><E"
        "lse><Statements><Action Name=\"MessageBox\"><Argument Name=\"Message\">The add-I"
        "n is not trusted curr"
End
Begin
    Comment ="_AXL:ently. Please click \"Enable Contents\" to trust the add-in and thus execut"
        "e the startup code normally.</Argument><Argument Name=\"Type\">Warning?</Argumen"
        "t><Argument Name=\"Title\">Add-in is not trusted</Argument></Action></Statements"
        "></Else></Condition"
End
Begin
    Comment ="_AXL:alBlock></Statements></UserInterfaceMacro>"
End
