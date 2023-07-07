Operation =1
Option =0
Where ="(((MSysNavPaneGroups.Name) Is Not Null) AND ((MSysNavPaneGroupCategories.Type)=4"
    "))"
Begin InputTables
    Name ="MSysNavPaneGroups"
    Name ="MSysNavPaneGroupToObjects"
    Name ="MSysObjects"
    Name ="MSysNavPaneGroupCategories"
End
Begin OutputColumns
    Alias ="CategoryName"
    Expression ="MSysNavPaneGroupCategories.Name"
    Alias ="CategoryPosition"
    Expression ="MSysNavPaneGroupCategories.Position"
    Alias ="CategoryFlags"
    Expression ="MSysNavPaneGroupCategories.Flags"
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
    Alias ="CategoryID"
    Expression ="MSysNavPaneGroupCategories.Id"
    Alias ="GroupID"
    Expression ="MSysNavPaneGroups.Id"
    Alias ="LinkID"
    Expression ="MSysNavPaneGroupToObjects.Id"
End
Begin Joins
    LeftTable ="MSysNavPaneGroupToObjects"
    RightTable ="MSysObjects"
    Expression ="MSysNavPaneGroupToObjects.ObjectID = MSysObjects.Id"
    Flag =2
    LeftTable ="MSysNavPaneGroupCategories"
    RightTable ="MSysNavPaneGroups"
    Expression ="MSysNavPaneGroupCategories.Id = MSysNavPaneGroups.GroupCategoryID"
    Flag =1
    LeftTable ="MSysNavPaneGroups"
    RightTable ="MSysNavPaneGroupToObjects"
    Expression ="MSysNavPaneGroups.Id = MSysNavPaneGroupToObjects.GroupID"
    Flag =2
End
Begin OrderBy
    Expression ="MSysNavPaneGroupCategories.Name"
    Flag =0
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
    Begin
        dbText "Name" ="CategoryName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CategoryPosition"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CategoryFlags"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LinkID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GroupID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CategoryID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1094
    Bottom =544
    Left =-1
    Top =-1
    Right =1078
    Bottom =397
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =251
        Top =42
        Right =440
        Bottom =234
        Top =0
        Name ="MSysNavPaneGroups"
        Name =""
    End
    Begin
        Left =513
        Top =40
        Right =752
        Bottom =247
        Top =0
        Name ="MSysNavPaneGroupToObjects"
        Name =""
    End
    Begin
        Left =853
        Top =40
        Right =997
        Bottom =376
        Top =0
        Name ="MSysObjects"
        Name =""
    End
    Begin
        Left =44
        Top =44
        Right =188
        Bottom =236
        Top =0
        Name ="MSysNavPaneGroupCategories"
        Name =""
    End
End
