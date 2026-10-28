# qryRegressionBracketedParenName

Regression canary for a Design View query whose source table has parentheses in its
bracketed name.

## Shape

`[tblCars (Archive)]` joined to `tblCarsModel`, with a `DesignLayout` in the `.json`.

## What broke

`clsQueryComposer.HasSubqueries` treated any input table whose name contained `(` (or
`SELECT `) as a derived table. Input table names are stored unbracketed, so
`[tblCars (Archive)]` became `tblCars (Archive)` and was taken for a subquery.
`IsDesignerCompatible` then returned False and the query was imported as SQL View,
silently dropping the layout that the `.json` carried. No warning is logged, because the
Design View path is never attempted.

## What must stay true

- A bracketed operand is a single identifier, whatever characters it contains.
- Derived tables (`(SELECT ...) AS x`, or a bare `FROM (SELECT ...)`) are still
  classified as subqueries and stay SQL View.
- The `import_path` check passes: Access stores a designer grid for the imported query.
