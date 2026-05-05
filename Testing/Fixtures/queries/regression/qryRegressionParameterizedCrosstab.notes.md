# qryRegressionParameterizedCrosstab

**Pins:** explicit `PARAMETERS` clauses on crosstab queries must survive
MSysQueries reconstruction.

Access stores this parameter as an Attribute 2 row with `Name1` holding the
parameter name and `Flag` holding the DAO type. The builder previously looked
only at `Expression`, which is null for this shape, and dropped the leading
`PARAMETERS [name] Text ( 255 );` line.
