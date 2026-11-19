# qryRegressionRecordLocksAppend

Regression canary for `RecordLocks` on a query imported in Design View.

## Shape

An append query with `RecordLocks = 2` and a `DesignLayout` in the `.json`.

## What broke

The same gap as `qryRegressionRecordLocksUpdate`: `RecordLocks` never reached the
`.qdef`, so a build lost it.

## What must stay true

- `RecordLocks` survives the import.
- The extra `dbByte "RecordLocks"` line does not make `LoadFromText` reject the
  Design View `.qdef`: the `import_path` check passes, so the query keeps its
  designer grid instead of falling back to SQL View.
