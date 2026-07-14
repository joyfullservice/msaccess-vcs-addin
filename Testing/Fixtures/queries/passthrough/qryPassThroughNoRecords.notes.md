# qryPassThroughNoRecords

**Pins:** Pass-through queries with Returns Records = No (MSysQueries Attribute 1
Flag = 10, `QueryType` 144 / `dbQSPTBulk`) must round-trip with verbatim action
SQL and `QueryProperties.ReturnsRecords = false`.

`ReturnsRecords` is not stored in `LvProp`; export derives it from the
Attribute 1 Flag. Import must emit `dbBoolean "ReturnsRecords" ="0"` so
`LoadFromText` restores the setting (GitHub issue #724).
