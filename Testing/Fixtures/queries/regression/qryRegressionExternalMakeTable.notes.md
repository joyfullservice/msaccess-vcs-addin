# qryRegressionExternalMakeTable

**Pins:** make-table queries that target an external Access database with
`IN '<path>'` must preserve the Attribute 1 `Name2` database path during SQL
reconstruction.

A production validation run exposed this on make-table queries where Access SQL
contained `SELECT ... INTO tblName IN 'C:\path\file.accdb' FROM ...`, but the
builder emitted only `INTO tblName`. That changes the destination database if
the query is executed.
