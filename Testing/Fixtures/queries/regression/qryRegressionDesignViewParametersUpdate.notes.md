# qryRegressionDesignViewParametersUpdate

Pins declared parameters on an **action** query taking the Design View import
path.

## Why this shape

Every other parameterized Design View fixture is a `SELECT`. An action query
differs in two ways that could plausibly have moved the parameters block:

- the header carries a different operation (`Operation =4` for `UPDATE` rather
  than `Operation =1`), and
- `Begin OutputColumns` holds `Name` / `Expression` pairs — the target column
  and the value assigned to it — instead of plain output expressions.

Native `Application.SaveAsText` output for a designer-built parameterized
`UPDATE` shows the parameters block in the same place regardless: immediately
after `Begin OutputColumns ... End`. The position is a property of the `.qdef`
grammar, not of the query type, and this fixture holds that finding in place.

A parameterized update is also a common real-world shape — "set this field for
the record I name" — so losing its parameters on rebuild would be conspicuous.

## What this fixture pins

- `QueryType` 48 with a `DesignLayout` forces the Design View path for an
  action query.
- The generated `.qdef` must emit `Operation =4`, the assignment pair in
  `Begin OutputColumns`, and `Begin Parameters` immediately after it.
- Round trip must preserve both parameters, one `Long` and one `Currency`.
