# qryRegressionExplicitAndAllFields

**Pins:** SELECT queries that combine explicit `table.*` output columns with
the Attribute 3 `AllFields` bit must reconstruct both projections.

A production validation run exposed a query stored as `SELECT tableAlias.*, *
FROM ...`. The builder was preserving the explicit `table.*` Attribute 6 row
but dropping the `*` implied by Option bit 1.
