# qryRegressionTopNDesignView

**Pins:** Design View import of `SELECT TOP N` queries must emit a
`RowCount ="N"` line in the `.qdef` after the `Option =` line.

A production build exposed three queries with TOP N + DesignLayout
that failed Design View import with "Object variable or With block
variable not set" and fell back to SQL View (losing layout). The
emitter wrote `Option =16` (the TOP flag) but omitted the
`RowCount` line that carries the actual count value. Without it,
Access's `LoadFromText` cannot reconstruct the query.

Specifically tests:

- `SELECT TOP 3` with DesignLayout metadata present.
- `OptionFlag: 16` (TOP bit) in the JSON companion.
- Design View `.qdef` must include `RowCount ="3"`.
- ORDER BY DESC in combination with TOP N.
