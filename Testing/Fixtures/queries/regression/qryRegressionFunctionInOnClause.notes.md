# qryRegressionFunctionInOnClause

**Pins:** Design View `.qdef` emitter must never emit a function-call operand as
`LeftTable` / `RightTable` when splitting a compound `ON` clause.

## The bug

For a condition like:

```sql
prior.OrderDate = DateAdd('yyyy', -1, cur.OrderDate)
```

`ExtractTableFromOnSide` used to isolate the right side, strip the trailing `)`,
find the first qualifying dot, and return the text before it — the literal
string `DateAdd('yyyy', -1, cur`. That garbage token was emitted as
`RightTable`. `Application.LoadFromText` accepted the `.qdef` silently; the
query then failed at runtime with DAO error 3080 ("Joined table not listed in
FROM clause").

The empty-string fallback in `EmitDesignViewQdef` did not help: extraction
returned a non-empty wrong value, so the parent join's tables were never used.

`ResolveConditionJoinTables` now ranks whole candidate pairs by whether they
cover every `InputTables` ref named in the condition. Here extraction resolves
only the left side (`prior`), so the parent pair `(cur, prior)` wins — it covers
both `prior` and `cur`. Resolving each side independently would instead yield
`(prior, cur)` and, for a single-table predicate, collapse to
`LeftTable = RightTable`; see `qryRegressionMultiCondJoin`.

## Expected emitter output

Both split conditions must carry `LeftTable ="cur"` / `RightTable ="prior"`
(parent-join orientation, so LEFT JOIN `Flag =2` stays correct), and both
refs must appear in `InputTables` (via `Alias`).

## How the harness catches it

The `qdef_joins` check verifies that every table referenced in a join
`Expression` is either that row's `LeftTable` or `RightTable`. With the old
extractor, `RightTable` was `DateAdd('yyyy', -1, cur`, which does not appear
as a qualifier in the expression (and is not in `InputTables`), so the check
fails. The unit test in `clsTestQueryComposerJoins` also asserts every join
ref exists in `InputTables`.

**See:** `docs/access-query-storage.md` § 5 / § 6, `DECISIONS.md` entry
`2026-08-10 — Function-call operands in ON clauses`.
