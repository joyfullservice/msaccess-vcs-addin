# qryRegressionMultiCondJoin

**Pins:** multi-condition `ON` clauses where the AND'd sub-conditions
reference different tables on each side.

A four-level LEFT JOIN chain in which two of the joins use
parenthesized multi-condition `ON` clauses, and one of the AND'd
conditions in the LEFT JOIN to `tblCarsPrice` references
`tblCarsModel.ModelID` -- a table joined two levels up, not the
immediate parent of the join.

The earlier `ExtractJoinLeftTable` / `ExtractJoinRightTable` logic
extracted table names by scanning *both* sides of the `ON` clause.
With three distinct tables in the `ON` clause it could pick the wrong
one, producing MSysQueries Attribute 7 rows where `Name1` and `Name2`
did not match the actual join inputs. The current
`TryExtractSimpleTable` + `OnClauseTableExcluding` strategy must
continue to handle this shape correctly; a regression here would
re-introduce the original multi-condition join corruption.
