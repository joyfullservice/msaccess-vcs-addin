# qryRegressionParametersCrosstabFixedPivot

Pins a declared parameter on a **crosstab** whose `PIVOT` clause names its column
headings explicitly, joined across two tables.

## Why the column headings are fixed

The `PIVOT` clause lists its headings (`In (1, 2, 3)`). Without them Access has
to *run* the query to discover the headings, so a parameterized crosstab prompts
for a parameter value the moment it is saved from the designer. Fixing the
headings keeps the fixture non-interactive, which matters because a prompt in an
automated run blocks until someone answers it.

Access normalizes the list to `In (1, 2, 3)` — with spaces after the commas —
when it stores the SQL, so the fixture is written that way. Authoring it as
`In (1,2,3)` produces a one-line diff on every run.

## Why this is a SQL View fixture

This started out carrying a `DesignLayout`, on the strength of a native
`Application.SaveAsText` capture showing that a designer-built parameterized
crosstab keeps its layout and puts `Begin Parameters` in the usual place
(immediately after `Begin OutputColumns ... End`). The add-in cannot reproduce
that yet: `clsQueryComposer.DecomposeSQL` deliberately marks every
`TRANSFORM`/`PIVOT` query as not designer-compatible, because the Design View
qdef generator does not emit the Attribute 6 aggregate/pivot fields Access
requires for `Operation =6`. A crosstab fixture with a `DesignLayout` therefore
imports through the SQL View path and silently loses the layout.

Rather than pin behaviour the add-in does not have, the fixture drops the
layout. Design View crosstab support is tracked as a known gap in
[docs/access-query-storage.md](../../../../docs/access-query-storage.md); when
it lands, restore the `DesignLayout` block and the `Begin Parameters` placement
assertion along with it.

## What this fixture pins

- A declared `Text` parameter survives a crosstab round trip.
- The `TRANSFORM` aggregate, the `GROUP BY` row heading, and the fixed `PIVOT`
  heading list all round-trip unchanged.
- The parameter is preserved even though the query is stored as SQL View, where
  the `PARAMETERS` clause is part of the SQL text rather than a structured block.
