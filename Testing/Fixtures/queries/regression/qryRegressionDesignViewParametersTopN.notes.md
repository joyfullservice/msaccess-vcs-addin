# qryRegressionDesignViewParametersTopN

Pins a parameterized `TOP n` query on the Design View import path.

## Why this shape

`TOP n` is the one shape where the `.qdef` header carries an extra element:
Access writes `Option =16` together with a `RowCount` line. Feeding it an
`Option =16` header with no matching `RowCount` produces a bare

```
Error encountered at line 3. Resource failure
```

which reports against the `Where` line rather than the option, so the real
cause is easy to misread. Pairing `TOP n` with a parameter keeps that header
combination under test alongside the parameters block.

The `Currency` parameter also covers the `dbCurrency` (flag 5) entry in the
type table, which no other Design View fixture exercises.

## What this fixture pins

- A `DesignLayout` in the companion `.json` forces `blnDesignView = True`, and
  `OptionFlag` 16 selects the `TOP n` header.
- The generated `.qdef` must emit `Option =16` with a `RowCount`, and place
  `Begin Parameters` immediately after `Begin OutputColumns ... End`, before
  `Begin OrderBy`.
- Round trip must preserve `TOP 5`, the `ORDER BY`, and the declared parameter
  with type `Currency`.
