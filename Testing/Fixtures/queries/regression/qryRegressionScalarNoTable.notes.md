# qryRegressionScalarNoTable

**Pins:** scalar SELECT queries with no input table must not gain Access's
empty-source `FROM ;` clause during reconstruction.

Access stores this shape with output-column rows but no Attribute 5 input table
rows. Empty SELECT queries and literal append queries legitimately use `FROM ;`;
scalar no-table SELECTs do not.
