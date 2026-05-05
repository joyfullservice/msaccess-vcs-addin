# qryRegressionUdfExpression

**Pins:** User-defined function calls as computed columns in a SELECT
expression, combined with a multi-condition ON clause.

No other fixture has function calls in SELECT expressions. The parser must
preserve `FunctionName([arg])` as an opaque expression without decomposing
the parentheses as join or subquery syntax.

Additional patterns exercised:
- Multi-condition ON using AND: `ON (A.Flags = B.Flags) AND (A.Type = B.Type)`
- Multiple computed columns with function calls and aliases

Derived from `qryMSysQueryObjects` in the MSysQueriesExamples v2 database.
Sanitized: `MSysObjects` replaced with `tblItems`, `tblSysObjectTypes` with
`tblItemTypes`, UDF names (`GetQueryLastSavedView`, `GetDateLastUpdated`)
replaced with generic `Func1`/`Func2`. Column metadata stripped.
