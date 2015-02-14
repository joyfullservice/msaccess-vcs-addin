Operation =1
Option =0
Where ="(((color_lookup.color)=\"red\"))"
Begin InputTables
    Name ="people"
    Name ="color_lookup"
End
Begin OutputColumns
    Expression ="people.full_name"
    Expression ="color_lookup.color"
End
Begin Joins
    LeftTable ="color_lookup"
    RightTable ="people"
    Expression ="color_lookup.id = people.favorite_color"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBinary "GUID" = Begin
    0x1d10bf75134eea4f83fa92cbf2ec3020
End
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="people.full_name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="color_lookup.color"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1705
    Bottom =927
    Left =-1
    Top =-1
    Right =1685
    Bottom =433
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="people"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="color_lookup"
        Name =""
    End
End
