# qryCarsCrosstab

A Design View crosstab with **two row headings** and an `ORDER BY`, over a single
source object.

## What this fixture pins

Two things the single-row-heading crosstab fixture cannot show:

- `GroupLevel` is a role marker, not a nesting depth. Both `Manufacturer` and
  `Year` are row headings and both carry `GroupLevel =2`; a second heading does
  not advance the level to 3. The `PIVOT` column heading is `1` and the
  `TRANSFORM` aggregate carries no marker at all.
- `Begin OrderBy` sits in its usual position on an `Operation =6` query, and its
  rows are ordinary `Flag =0` entries with no crosstab annotation.

It also covers the plain case of a `PIVOT` with no fixed heading list, where the
column-heading expression and its matching `Groups` row are identical.

## Where the layout came from

`DesignLayout` is a real capture. The query was rebuilt against a scratch
`qryCars`, saved from the designer, and its `MSysObjects.LvExtra` blob decoded
with `clsLvExtraParser`. `ColumnsShown = 559` matches the parameterized crosstab
fixture, confirming that `ORDER BY` does not change which grid rows are shown.

Unlike a parameterized crosstab, this one has no parameters, so Access can run
it to discover the column headings without prompting — which is why it needs no
`In (...)` list to be authored in the designer.
