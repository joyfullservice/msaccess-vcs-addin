# qryRegressionTopNGroupByDesignView

**Pins:** Design View import of `SELECT TOP N ... GROUP BY` queries
must emit both the `RowCount` line and the `Groups` block.

A production build exposed a query combining TOP 1 with GROUP BY
and INNER JOIN that failed Design View import. The `.qdef` was
missing the `RowCount ="1"` line after `Option =16`.

Specifically tests:

- `SELECT TOP 1` with GROUP BY and DesignLayout.
- INNER JOIN between two tables.
- `ColumnsShown: 543` (539 base + 4 for GROUP BY totals row).
- Design View `.qdef` must include `RowCount ="1"` and `Groups`.
