# qryPassThroughReturnsRecords

**Pins:** Pass-through queries with Returns Records = Yes (MSysQueries Attribute 1
Flag = 8, `QueryType` 112 / `dbQSQLPassThrough`) must omit non-default
`ReturnsRecords` from export JSON while preserving verbatim SELECT SQL.

Pairs with `qryPassThroughNoRecords` (Flag 10) to cover both pass-through
Attribute 1 flag shapes.
