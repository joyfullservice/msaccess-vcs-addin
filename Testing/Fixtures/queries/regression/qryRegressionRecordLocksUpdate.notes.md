# qryRegressionRecordLocksUpdate

Regression canary for `RecordLocks` on an action query imported in SQL View.

## Shape

An UPDATE query with `RecordLocks = 2` (Edited Record) and `FailOnError = True`
in its `QueryProperties`.

## What broke

`clsQueryComposer.EmitAllProperties` skipped `RecordLocks` as LvProp-only, so the
`.qdef` handed to `LoadFromText` never carried it and a build lost the value. The
export of the built query then had no `RecordLocks` (issue #813).

## What must stay true

- The `.qdef` carries `dbByte "RecordLocks"`. `SaveAsText` does not write that
  line, but `LoadFromText` accepts it.
- The re-export keeps `RecordLocks` after `FailOnError`, in the order the
  composer emits them, so Pass 1 matches the fixture.
