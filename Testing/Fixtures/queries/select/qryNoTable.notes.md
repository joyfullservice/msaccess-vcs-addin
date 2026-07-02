# qryNoTable

**Pins:** SELECT with an empty column list and an empty FROM clause (`FROM ;`).

Access allows queries with no source table and no projected columns, producing
the `SELECT FROM ;` pattern. This complements `qryRegressionAppendValues`
(which also has `FROM ;`, but inside an INSERT INTO context). Here the bare
`FROM ;` appears in a plain SELECT, exercising the same
`TrimTrailingSemicolon` / `FindTopLevelKeyword("FROM ")` interaction from a
different entry point.
