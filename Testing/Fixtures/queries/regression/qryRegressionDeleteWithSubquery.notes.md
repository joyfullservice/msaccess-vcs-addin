# qryRegressionDeleteWithSubquery

**Pins:** DELETE queries whose criteria contain a scalar `Exists (SELECT ...)`
subquery must preserve the subquery expression in the SQL text and generated
qdef.

The query is never executed by the round-trip harness; it only verifies
definition import/export fidelity.
