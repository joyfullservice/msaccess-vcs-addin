# qryRegressionCrossTableOn

**Pins:** cross-table multi-condition `ON` clauses where one logical JOIN is
stored by Access as multiple MSysQueries Attribute 7 rows with different table
pairs.

This shape mirrors a production failure where Access stores one JOIN from an
outer table into a nested subtree, with one `ON` predicate referencing one table
inside the subtree and another predicate referencing a different table inside
the same subtree.

The SQL builder must merge those secondary Attribute 7 rows back into the
same logical `ON (...) AND (...)` clause. If `ConsolidateJoins` treats the
secondary predicate as a standalone join, `BuildFromClause` can encounter
"both join tables already in FROM clause" and silently lose a predicate.
