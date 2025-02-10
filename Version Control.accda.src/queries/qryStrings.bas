Operation =1
Option =0
Begin InputTables
    Name ="tblLanguages"
    Name ="tblStrings"
End
Begin OutputColumns
    Alias ="LanguageID"
    Expression ="tblLanguages.ID"
    Expression ="tblStrings.ID"
    Expression ="tblStrings.msgid"
    Expression ="tblStrings.Context"
    Expression ="tblStrings.Reference"
    Expression ="tblStrings.Comments"
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
        dbText "Name" ="tblStrings.Comments"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LanguageID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblStrings.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblStrings.msgid"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3765"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblStrings.Context"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblStrings.Reference"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1705
    Bottom =848
    Left =-1
    Top =-1
    Right =1689
    Bottom =535
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =145
        Top =91
        Right =291
        Bottom =265
        Top =0
        Name ="tblLanguages"
        Name =""
    End
    Begin
        Left =349
        Top =91
        Right =503
        Bottom =263
        Top =0
        Name ="tblStrings"
        Name =""
    End
End
