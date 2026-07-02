# qryRegressionCrossTableOn

**Pins:** cross-table multi-condition `ON` clauses where one logical JOIN is
stored by Access as multiple MSysQueries Attribute 7 rows with different table
pairs.

This shape mirrors a production failure where Access stores one JOIN from an
outer table into a nested subtree, with one `ON` predicate referencing one table
inside the subtree and another predicate referencing a different table inside
the same subtree.

## SQL reconstruction (ConsolidateJoins)

The SQL builder must merge those secondary Attribute 7 rows back into the
same logical `ON (...) AND (...)` clause. If `ConsolidateJoins` treats the
secondary predicate as a standalone join, `BuildFromClause` can encounter
"both join tables already in FROM clause" and silently lose a predicate.

## Design View qdef emission (EmitDesignViewQdef)

When the emitter splits a compound `ON` clause back into separate Design View
`.qdef` rows, each row must carry the correct `LeftTable`/`RightTable` pair
extracted from that specific condition expression — not the parent join's table
pair. Access stores each Attribute 7 row with the specific table pair that the
condition references (e.g., `ON (A.x = B.x) AND (A.x = C.x)` becomes two rows:
`A`-`B` and `A`-`C`).

Reusing the parent join's tables for all split conditions causes
`Application.LoadFromText` to create ambiguous internal references. When one of
the referenced tables is also used by a saved query in another condition (e.g.,
table `T` is referenced directly in one condition and also appears inside
subquery `Q` in another), the wrong `RightTable` causes DAO error 3082 ("JOIN
operation refers to a field that is not in one of the joined tables") at query
execution time.

**See:** `docs/access-query-storage.md` § 6, `DECISIONS.md` entry
`2026-05-07 — Cross-table ON condition LeftTable/RightTable in Design View qdef`.
