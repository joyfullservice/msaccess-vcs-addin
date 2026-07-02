# qryRegressionFindDuplicates

**Pins:** the Access "Find Duplicates" wizard pattern -- a SELECT with
a self-referencing scalar subquery in
`WHERE col In (SELECT col FROM [tbl] AS Tmp ...)`.

This shape triggered "Resource failure" Design View import errors on
four queries during a full-database build (errors at line 2, line 3,
line 3, and line 37 of the rebuilt .qdef respectively). Affected
queries fell back to SQL View, losing their designer layout. The
wizard generates this exact form whenever the user runs
Wizard > Find Duplicates, so the pattern occurs in many real
databases.

Specifically tests:

- Aliased self-reference inside a subquery (`[tblCars] AS Tmp`).
- Bracketed identifiers (`[Manufacturer]`).
- Scalar subquery as a value source for `IN`.
- `Count(*) > 1` aggregate filter inside `HAVING`.

If this fixture starts failing the import (or the .qdef the importer
generates becomes unparseable), the design-view fallback regression
has returned.
