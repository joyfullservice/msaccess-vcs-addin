# qryRegressionRecordLocksSelect

Regression canary for `RecordLocks` on a select query imported in Design View,
next to a non-default datasheet property.

## Shape

A select query with `RecordLocks = 2` and `DefaultView = 0`, and a `DesignLayout`
in the `.json`.

## What broke

The same gap as `qryRegressionRecordLocksUpdate`: `RecordLocks` never reached the
`.qdef`, so a build lost it.

## What must stay true

- `RecordLocks` survives the import, and the extra line does not make
  `LoadFromText` reject the Design View `.qdef` (`import_path` passes).
- The `.json` lists `RecordLocks` after `DefaultView`, the order in which the
  composer emits them, so Pass 1 matches the fixture.
