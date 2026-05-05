# qryRegressionUnionOrderBy

**Pins:** UNION queries must preserve global Attribute 11 `ORDER BY` rows.

A production validation run found several UNION/UNION ALL queries where the
builder emitted all branches but terminated before appending the final sort.
That changes observable ordering even when the row set is identical.
