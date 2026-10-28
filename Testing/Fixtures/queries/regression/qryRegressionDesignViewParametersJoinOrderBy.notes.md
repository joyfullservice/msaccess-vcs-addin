# qryRegressionDesignViewParametersJoinOrderBy

Pins the **position** of the `Begin Parameters` block in a Design View `.qdef`.

## The bug

`qryRegressionDesignViewParameters` and `qryRegressionDesignViewParameterTypes`
proved the emitter writes a `Begin Parameters` block, but both are single-table
queries with no join and no `ORDER BY`. In that shape the parameters block lands
immediately after `Begin OutputColumns ... End`, which is where native
`SaveAsText` puts it — so they round-trip cleanly.

The emitter originally wrote the block *after* the `Joins`, `OrderBy` and
`Groups` blocks (just before the query properties). `Application.LoadFromText`
rejects a `Begin Parameters` block in that position with:

```
Error encountered at line N.
Expected: End of file.  Found: Parameters.
```

Any parameterized query that also had a join or an `ORDER BY` therefore failed
the Design View import and fell back to SQL View — losing its stored grid
layout. Production validation hit this on four queries at once; the two
single-table fixtures never exercised it.

## What this fixture pins

- A `DesignLayout` in the companion `.json` (two tables) forces
  `blnDesignView = True`.
- The `.sql` combines **a join** (`tblCustomers INNER JOIN tblOrders`), an
  **`ORDER BY`**, and **two parameters** (`[Min Qty]` Long, `StatusFilter`
  Short — one bracketed, one not).
- Round-trip (import → export) must preserve both parameters *and* the layout,
  and the generated `.qdef` must place the `Begin Parameters` block immediately
  after `Begin OutputColumns ... End`, before `Begin Joins`.

If `EmitParameters` is moved back after the Joins/OrderBy/Groups blocks, this
fixture's import fails and the round trip drops to the SQL View fallback.
