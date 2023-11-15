Operation =1
Option =0
Begin InputTables
End
Begin OutputColumns
    Alias ="FormControl"
    Expression ="[Forms]![frmColors]![Text18]"
    Alias ="TestExpression"
    Expression ="IIf([Forms]![frmVCSInstall]![chkUseRibbon],Eval(\"True\"),False)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="FormControl"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1590"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="TestExpression"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1368
    Bottom =856
    Left =-1
    Top =-1
    Right =1352
    Bottom =577
    Left =0
    Top =0
    ColumnsShown =539
End
