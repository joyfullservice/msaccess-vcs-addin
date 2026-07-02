# qryRegressionFromSubquery

**Pins:** a derived table (subquery) in the FROM clause whose internal
alias is Access's `%$##@_Alias` placeholder. This is the shape Access
produces when the query designer wraps an unnamed subquery in the FROM.

The minimal failing artifact looks like:

```sql
SELECT [%$##@_Alias].Col FROM (SELECT ... FROM ...) AS [%$##@_Alias];
```

In MSysQueries this is stored as:

- **Attribute 5** (Input Table): `Name1` = the entire inner SELECT text,
  `Name2` = `%$##@_Alias`
- **Attribute 6** (Output Columns): one row per outer projection, each
  `Expression` referencing `[%$##@_Alias].<col>`

Specifically tests:

- A FROM-clause derived table (subquery) -- `clsQueryComposer.ReconstructSQL`
  must wrap the InputTable's `Name1` in `(...) AS [<alias>]` rather than
  emit `[<entire-select>] AS [<alias>]` (the symptom of the
  pre-fix `BracketIfNeeded`-as-table-name behavior).
- The literal `[%$##@_Alias]` token in the outer projection -- the
  formatter must not strip or rewrite it.
- The composer's `IsDesignerCompatible` gate (`HasSubqueries()` returning
  True) -- export must emit SQL View qdef, not Design View, because
  Access's `LoadFromText` rejects the legacy `InputTables.Name = "<SELECT>"`
  / `Alias = "%$##@_Alias"` shape with "Resource failure" at the Alias
  line.
- The inner subquery preserves: an INNER JOIN, multi-condition WHERE,
  string concatenation, an `IN (SELECT ...)` nested subquery, an
  `ORDER BY`, and an outer-side function call wrapping a derived-alias
  column.

## Real-world origin

User-reported reproduction in 4.x format showed
`Microsoft Access encountered an error while importing the object ...
Error encountered at line 12. Resource failure` when building from
source. Line 12 of the legacy `.bas` was the
`Alias ="%$##@_Alias"` row. The new `.sql`/`.json` pipeline routes
queries with subqueries through SQL View qdef
(`dbMemo "SQL" = "..."`) instead, which `LoadFromText` accepts -- but
only if the export side correctly preserves the `(SELECT ...) AS [alias]`
shape rather than mangling the InputTable name.

If this fixture starts failing with malformed FROM-clause SQL on
re-export (e.g. the inner SELECT appearing bracketed as if it were a
literal table name), the FROM-subquery handling regression has returned.
