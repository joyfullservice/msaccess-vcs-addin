Operation =1
Option =0
Begin InputTables
    Name ="tblStrings"
    Name ="tblTranslation"
End
Begin OutputColumns
    Expression ="tblStrings.ID"
    Alias ="Key"
    Expression ="[Context] & \"|\" & [msgid]"
    Expression ="tblStrings.msgid"
    Expression ="tblStrings.Context"
    Expression ="tblStrings.Comments"
    Expression ="tblTranslation.Translation"
    Expression ="tblTranslation.Language"
End
Begin Joins
    LeftTable ="tblStrings"
    RightTable ="tblTranslation"
    Expression ="tblStrings.ID = tblTranslation.StringID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblTranslation.Translation"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="7215"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Key"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3525"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblStrings.ID"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="870"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblStrings.Comments"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3510"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblTranslation.Language"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblStrings.msgid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblStrings.Context"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1705
    Bottom =856
    Left =-1
    Top =-1
    Right =1299
    Bottom =526
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =162
        Top =102
        Right =306
        Bottom =269
        Top =0
        Name ="tblStrings"
        Name =""
    End
    Begin
        Left =438
        Top =125
        Right =582
        Bottom =269
        Top =0
        Name ="tblTranslation"
        Name =""
    End
End
