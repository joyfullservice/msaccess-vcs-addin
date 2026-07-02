# qryRegressionDualStarDesignView

**Pins:** Design View import of `SELECT table.*, *` queries must
represent the bare `*` as Option bit 1 (Output All Fields), not as
an explicit `Expression ="*"` in OutputColumns.

A production build exposed two queries combining `table.*` with `*`
that failed Design View import. The parser was creating a second
OutputColumn for the bare `*`, but Access stores this pattern as
Option bit 1 plus a single Attribute 6 row for `table.*` only.
The spurious `Expression ="*"` in OutputColumns caused
`LoadFromText` to reject the `.qdef`.

Specifically tests:

- `SELECT tblCars.*, *` with DesignLayout metadata present.
- `OptionFlag: 1` (Output All Fields) in the JSON companion.
- The bare `*` must become Option bit 1, not an OutputColumn.
- WHERE clause coexists with dual-star projection.

Companion to `qryRegressionExplicitAndAllFields` which tests the
same SQL shape in SQL View (no DesignLayout).
