# qryRegressionNoFormatUnion

Regression canary for `NoFormat` on a union query.

## Shape

A UNION query with `NoFormat = False` in its `QueryProperties`.

## What broke

`NoFormat` fell through to the `Case Else` of `EmitAllProperties`, so the `.qdef`
never carried it and a build lost the value.

## What must stay true

- The `.qdef` carries `dbBoolean "NoFormat" ="0"`: the JSON `false` is normalized
  to the Access boolean string, as for the other boolean query properties.
- The re-export keeps `NoFormat`.
