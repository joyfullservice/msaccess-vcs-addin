# qryRegressionSpacedTableNameJoin

Regression canary for bracketed multi-word table names used as join operands.

## Shape

`SELECT ... FROM [Name With Spaces] INNER JOIN [Other Name] ON ...`

## What broke

`TryExtractSimpleTable` un-bracketed the operand and split on the first space, producing
truncated join references (`Car` instead of `Car Models`). The malformed qdef
was accepted by `LoadFromText` but corrupted `MSysQueries`, and the next export emitted
un-importable SQL.

## What must stay true

- Join `LeftTable`/`RightTable` values in the generated qdef must match `InputTables` names.
- Export and import paths must preserve the full bracketed table names with no warnings.
