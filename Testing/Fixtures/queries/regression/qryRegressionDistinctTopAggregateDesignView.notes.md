# qryRegressionDistinctTopAggregateDesignView

**Pins:** Design View import of `SELECT DISTINCT TOP N` queries with
aggregate expressions in both SELECT and ORDER BY must emit the
`RowCount` line and handle aggregate columns correctly.

A production build exposed a query combining DISTINCT + TOP 3 with
`Max()` aggregate in both SELECT and ORDER BY (no explicit GROUP BY)
that failed Design View import. `OptionFlag: 18` combines
DISTINCT (bit 1) and TOP (bit 4). The `.qdef` was missing
`RowCount ="3"`.

Specifically tests:

- `SELECT DISTINCT TOP 3` with DesignLayout.
- `OptionFlag: 18` (DISTINCT + TOP) in JSON.
- Aggregate expression `Max(Year([Date])-1)` in SELECT and ORDER BY.
- `ColumnsShown: 543` (GROUP BY totals row visible).
- Design View `.qdef` must include `RowCount ="3"`.
