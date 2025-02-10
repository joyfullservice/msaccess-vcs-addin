Operation =1
Option =0
Begin InputTables
    Name ="tblInternal"
End
Begin OutputColumns
    Alias ="DatabaseFile"
    Expression ="GetDatabaseFileName()"
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
        dbText "Name" ="DatabaseFile"
        dbInteger "ColumnWidth" ="8700"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1305
    Bottom =848
    Left =-1
    Top =-1
    Right =1289
    Bottom =586
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =105
        Top =99
        Right =277
        Bottom =318
        Top =0
        Name ="tblInternal"
        Name =""
    End
End
