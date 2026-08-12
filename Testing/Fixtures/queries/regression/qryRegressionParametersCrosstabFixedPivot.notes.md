# qryRegressionParametersCrosstabFixedPivot

Pins a declared parameter on a **crosstab** whose `PIVOT` clause names its column
headings explicitly, joined across two tables.

## Why the column headings are fixed

This is an authoring constraint, not an import one. The `PIVOT` clause lists its
headings (`In (1, 2, 3)`) so the fixture can be *built* in the Access designer:
without them Access has to run the query to discover the headings, and a
parameterized crosstab then prompts for a parameter value the moment it is
saved. `DoCmd.SetParameter` does not suppress that prompt. The add-in's own
import path is unaffected either way — `LoadFromText` never executes the query —
but the layout below had to come from a designer save, so the fixture has to be
a shape the designer can save unattended.

Access normalizes the list to `In (1, 2, 3)` — with spaces after the commas —
when it stores the SQL, so the fixture is written that way. Authoring it as
`In (1,2,3)` produces a one-line diff on every run.

## Where the layout came from

`DesignLayout` is a real capture, not a plausible-looking hand edit: the query
was rebuilt against scratch copies of `tblParamSample` and `tblParamDetail`,
saved from the designer, and its `MSysObjects.LvExtra` blob decoded with
`clsLvExtraParser`. `ColumnsShown = 559` is the crosstab grid — the ordinary
select value is 539 and a plain `GROUP BY` is 543, so a crosstab is not simply
"543 plus a bit".

## What this fixture pins

- A declared `Text` parameter survives a crosstab round trip.
- The `TRANSFORM` aggregate, the `GROUP BY` row heading, and the fixed `PIVOT`
  heading list all round-trip unchanged.
- `Begin Parameters` keeps its usual position on an `Operation =6` query,
  immediately after `Begin OutputColumns ... End`.
- The designer grid survives import, which the `import_path` check asserts by
  looking for a non-null `LvExtra` afterwards.
