# qryRegressionCrossTableOnEmitter

**Pins:** correct `LeftTable`/`RightTable` assignment in the Design View `.qdef`
emitter when a compound `ON` clause has individual conditions referencing
different table pairs.

## The bug

When `EmitDesignViewQdef` splits a compound `ON` clause (e.g.,
`ON (A.x = B.x) AND (A.y = C.y)`) into separate `.qdef` join rows, each
row must carry the table pair from *its own condition*, not the parent join's
tables. Before the fix, the emitter reused the parent join's `rightTable`
for all split conditions, producing:

```
LeftTable ="tblOrders"
RightTable ="tblProducts"         <-- wrong for condition 2
Expression ="tblOrders.CategoryID = tblCategories.CategoryID"
```

The correct output is `RightTable ="tblCategories"` for condition 2.

## Why this matters

`Application.LoadFromText` silently accepts the wrong `RightTable`.
The error only surfaces at query execution time as DAO error 3082 ("JOIN
operation refers to a field that is not in one of the joined tables") when
the tables actually exist in the database and one of them is also used
inside a saved subquery referenced in another condition.

## How the harness catches it

The `qdef_joins` validation check (added alongside this fixture) generates
the Design View `.qdef` for the fixture SQL and verifies that each join
row's `LeftTable`/`RightTable` appears in its `Expression`. With the old
emitter, condition 2 would have `RightTable ="tblProducts"` but Expression
`tblOrders.CategoryID = tblCategories.CategoryID` — `tblProducts` does not
appear in the expression, so the check fails.

**See:** `docs/access-query-storage.md` § 6, `DECISIONS.md` entry
`2026-05-07 — Cross-table ON condition LeftTable/RightTable in Design View qdef`.
