# qryRegressionDesignViewParametersGroupBy

Pins the position of the `Begin Parameters` block relative to `Begin Groups`.

## Why this shape

`qryRegressionDesignViewParametersJoinOrderBy` proves the block must precede
`Begin Joins` and `Begin OrderBy`. `Begin Groups` is emitted *after* both of
those, so a totals query is the only shape that pins the block against it.

The constraint was established by feeding `Application.LoadFromText` the same
Design View `.qdef` with the parameters block in each candidate position. Only
one position loads:

| Position of `Begin Parameters`     | Result                                    |
| ---------------------------------- | ----------------------------------------- |
| After `Begin OutputColumns ... End` | Loads; Access re-emits it in place        |
| After `Begin Joins ... End`         | `Expected: End of file.  Found: Parameters.` |
| After `Begin OrderBy ... End`       | `Expected: End of file.  Found: Parameters.` |
| After `Begin Groups ... End`        | `Expected: 'End'.  Found: Parameters.`    |
| Before `Begin InputTables`          | `Expected: End of file.  Found: InputTables.` |

Native `Application.SaveAsText` output for a designer-built totals query agrees:
the block sits immediately after the output columns, ahead of `Begin OrderBy`
and `Begin Groups`.

## What this fixture pins

- A `DesignLayout` in the companion `.json` forces `blnDesignView = True`.
- The `.sql` combines an aggregate (`GROUP BY` plus `Count`), an `ORDER BY`, and
  one typed parameter.
- The generated `.qdef` must place `Begin Parameters` immediately after
  `Begin OutputColumns ... End`, before both `Begin OrderBy` and `Begin Groups`.

If the block moves after the groups, the import fails outright and the round
trip silently drops to the SQL View fallback, losing the stored grid layout.
