# qryRegressionJoinPredicatePlacement

**Pins:** cross-subtree `ON` predicates must stay attached to the join that
connects the root table to the nested subtree.

A production validation run exposed this when a predicate joining the root
table to a table inside the nested subtree was moved onto a later join to an
unrelated table that merely shared the root table. The generated SQL was no
longer parseable by Access ODBC (`JOIN expression not supported`).
