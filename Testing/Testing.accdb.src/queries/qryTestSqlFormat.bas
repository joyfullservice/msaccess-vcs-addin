dbMemo "SQL" ="SELECT MSysQueries.Attribute, MSysQueries.Flag\015\012FROM MSysQueries\015\012WH"
    "ERE (((MSysQueries.Flag) In (SELECT\015\012           Flag\015\012         from\015"
    "\012           [MSysQueries]\015\012         where\015\012           flag = 0\015"
    "\012         )));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
