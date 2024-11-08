Operation =1
Option =0
Begin InputTables
    Name ="tblTranslation"
    Name ="qryStrings"
End
Begin OutputColumns
    Expression ="qryStrings.ID"
    Expression ="qryStrings.msgid"
    Expression ="qryStrings.Context"
    Expression ="qryStrings.Comments"
    Expression ="tblTranslation.Translation"
    Alias ="Lang"
    Expression ="qryStrings.LanguageID"
    Expression ="qryStrings.Reference"
    Alias ="SortRank"
    Expression ="IIf([msgid]=\"\",1,2)"
    Alias ="Key"
    Expression ="[Context] & \"|\" & [msgid]"
End
Begin Joins
    LeftTable ="qryStrings"
    RightTable ="tblTranslation"
    Expression ="qryStrings.ID = tblTranslation.StringID"
    Flag =2
    LeftTable ="qryStrings"
    RightTable ="tblTranslation"
    Expression ="qryStrings.LanguageID = tblTranslation.Language"
    Flag =2
End
Begin OrderBy
    Expression ="IIf([msgid]=\"\",1,2)"
    Flag =0
    Expression ="qryStrings.msgid"
    Flag =0
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
        dbText "Name" ="Lang"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SortRank"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryStrings.ID"
        dbInteger "ColumnWidth" ="870"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryStrings.msgid"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryStrings.Context"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryStrings.Comments"
        dbInteger "ColumnWidth" ="3510"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryStrings.Reference"
        dbInteger "ColumnWidth" ="4290"
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
    Bottom =603
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =453
        Top =146
        Right =597
        Bottom =290
        Top =0
        Name ="tblTranslation"
        Name =""
    End
    Begin
        Left =199
        Top =112
        Right =347
        Bottom =295
        Top =0
        Name ="qryStrings"
        Name =""
    End
End
