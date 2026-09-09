# qryRegressionAppendDistinctTopPercent

**Pins:** `TOP n PERCENT` inside an `INSERT ... SELECT` is a modifier, not the
first output column.

`ParseInsertQuery` stripped `DISTINCT` / `DISTINCTROW` from the SELECT portion
but never looked for the `TOP` clause that follows it, so `TOP 5 PERCENT` stayed
in the field list and was parsed as an output column expression. The append
query then re-exported with no TOP bits and a phantom column.

The companion `OptionFlag: 50` (2 DISTINCT + 16 TOP + 32 PERCENT) agrees with
the `.sql`, which is the only combination the round-trip invariant accepts. The
`DesignLayout` keeps this on the Design View import path, where the parsed
modifiers have to reach `Option` and `RowCount` in the `.qdef` rather than
being re-parsed by Access from a SQL View memo.
