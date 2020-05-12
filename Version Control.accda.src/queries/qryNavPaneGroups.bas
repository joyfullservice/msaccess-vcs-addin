Operation =1
Option =0
Where ="(((MSysNavPaneGroups.Name) Is Not Null) AND ((MSysNavPaneGroups.GroupCategoryID)"
    "=3))"
Begin InputTables
    Name ="MSysNavPaneGroups"
    Name ="MSysNavPaneGroupToObjects"
    Name ="MSysObjects"
End
Begin OutputColumns
    Alias ="GroupName"
    Expression ="MSysNavPaneGroups.Name"
    Alias ="GroupFlags"
    Expression ="MSysNavPaneGroups.Flags"
    Alias ="GroupPosition"
    Expression ="MSysNavPaneGroups.Position"
    Alias ="ObjectType"
    Expression ="MSysObjects.Type"
    Alias ="ObjectName"
    Expression ="MSysObjects.Name"
    Alias ="ObjectFlags"
    Expression ="MSysNavPaneGroupToObjects.Flags"
    Alias ="ObjectIcon"
    Expression ="MSysNavPaneGroupToObjects.Icon"
    Alias ="ObjectPosition"
    Expression ="MSysNavPaneGroupToObjects.Position"
    Alias ="NameInGroup"
    Expression ="MSysNavPaneGroupToObjects.Name"
End
Begin Joins
    LeftTable ="MSysNavPaneGroupToObjects"
    RightTable ="MSysObjects"
    Expression ="MSysNavPaneGroupToObjects.ObjectID = MSysObjects.Id"
    Flag =2
    LeftTable ="MSysNavPaneGroups"
    RightTable ="MSysNavPaneGroupToObjects"
    Expression ="MSysNavPaneGroups.Id = MSysNavPaneGroupToObjects.GroupID"
    Flag =2
End
Begin OrderBy
    Expression ="MSysNavPaneGroups.Name"
    Flag =0
    Expression ="MSysObjects.Type"
    Flag =0
    Expression ="MSysObjects.Name"
    Flag =0
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
        dbText "Name" ="GroupName"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2085"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="ObjectName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ObjectType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupFlags"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ObjectFlags"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ObjectIcon"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ObjectPosition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupPosition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NameInGroup"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =646
    Bottom =800
    Left =-1
    Top =-1
    Right =630
    Bottom =504
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="MSysNavPaneGroups"
        Name =""
    End
    Begin
        Left =238
        Top =11
        Right =389
        Bottom =187
        Top =0
        Name ="MSysNavPaneGroupToObjects"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =210
        Top =0
        Name ="MSysObjects"
        Name =""
    End
End
