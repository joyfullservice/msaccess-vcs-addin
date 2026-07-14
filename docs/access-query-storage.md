# Access Query Storage Reference

In-repo reference for how Microsoft Access stores queries internally and what
this add-in's query parser (`clsQueryComposer` + `clsDbQuery`) currently
handles, deliberately doesn't handle, or has had to work around.

## Provenance and attribution

Three sources contribute to this document:

- **Primary source for the field reference (sections 2 and 3):** Colin
  Riddington, Mendip Data Systems —
  [Part 1: The MSysQueries Table](https://www.isladogs.co.uk/explaining-queries/index.html)
  (first published 3 Feb 2019, last updated 7 Aug 2022) and
  [Part 2: Design View vs SQL View](https://www.isladogs.co.uk/explaining-queries-2/index.html)
  (8 Aug 2022, last updated 1 Sept 2022).
- **Worked-example queries:** also Colin Riddington — provided as an
  example download alongside the articles. A subset of those queries
  has been ported into this add-in's fixture corpus
  (`Testing/Fixtures/queries/`) to prove round-trip idempotence on the
  shapes Colin documents.
- **Secondary source / packaging:** the `MSysQueriesExamples` companion
  repo (Adam Waller) is an internally-maintained package built around
  Colin's example download — it organizes the queries by topic in a
  live test database and adds the April 2026 Lv/MSysObjects
  binary-format addendum. Lives at `C:\Repos\MSysQueriesExamples\` (not
  vendored into this repo); see `docs/how-access-stores-queries.md`
  there for the exhaustive treatment.
- **Unique to this repo (sections 4–6):** empirical findings from
  running the round-trip fixture corpus and the `clsQueryComposer`
  test runs, including `Application.LoadFromText` /
  `Application.SaveAsText` asymmetries that are not in the upstream
  documentation.

## 1. Scope and how to read this doc

**Covers:**

- The MSysQueries fields the parser actually reads from (Attributes 1, 3, 5,
  6, 7, 8, 9, 10, 11, plus the undocumented `Order` byte sequence) and the
  MSysObjects `LvProp` / `LvExtra` blobs the exporter consumes.
- The Design-View-vs-SQL-View arbitration logic the importer uses to
  reproduce a query in its last-saved view.
- A complete table of which query shapes have a fixture and which don't.
- Findings unique to our import pipeline (`LoadFromText` / `SaveAsText`
  asymmetries, the legacy `ForceImportOriginalQuerySQL` safety net).

**Does not cover:**

- Exhaustive Access query trivia. Use `MSysQueriesExamples` /
  `how-access-stores-queries.md` or the isladogs articles for that.
- Why specific decisions were made. That's `DECISIONS.md` at the repo root.
  This doc captures *what is true today*; `DECISIONS.md` captures *why we
  chose it and what alternatives were ruled out*.
- One-bug-one-shape regression context. That's the per-fixture `.notes.md`
  files alongside `Testing/Fixtures/queries/regression/*.sql`.

## 2. MSysQueries field reference

Every Access query is stored as a row sequence in `MSysQueries`, joined to a
single `MSysObjects` parent row by `ObjectId`. Each row carries an
`Attribute`, `Flag`, `Expression`, `Name1`, and `Name2`, plus an
undocumented `Order` blob that establishes the row sequence within an
`Attribute` group. Read with `ORDER BY Attribute, [Order]` to match the
order `Application.SaveAsText` emits.

### Attribute 1 — query type

`Flag` carries the query kind. The full set is in the
[Attribute 1 Quick Reference](https://www.isladogs.co.uk/explaining-queries/index.html).
The parser handles every value below:

| Flag | Query type                | Where Name1/Name2 carry data                           |
|------|---------------------------|--------------------------------------------------------|
| 1    | SELECT                    | —                                                      |
| 3    | Append (INSERT INTO)      | Name1 = target table; Name2 = external DB path if any  |
| 4    | UPDATE                    | —                                                      |
| 5    | DELETE                    | —                                                      |
| 6    | TRANSFORM (Crosstab)      | —                                                      |
| 7    | Data Definition (DDL)     | Expression = full DDL statement                        |
| 8    | Pass-through (returns records) | Expression = SQL command; Name1 = ODBC connect string  |
| 10   | Pass-through (no records)        | Expression = SQL command; Name1 = ODBC connect string  |
| 9    | UNION                     | —                                                      |

Flag 2 (Make-Table / `SELECT INTO`) is in the upstream reference but has no
fixture in this repo (see § 5).

### Attribute 3 — query options (bitmask)

The full set:

| Flag | Option                                  |
|------|------------------------------------------|
| 0    | None (default — omitted in SQL View)    |
| 1    | Output All Fields / UNION ALL / TEMP    |
| 2    | DISTINCT (Unique Values)                |
| 3    | UNION                                   |
| 4    | WITH OWNERACCESS OPTION                 |
| 8    | DISTINCTROW                             |
| 9    | TEMP query                              |
| 16   | TOP *n*                                 |
| 18   | DISTINCT + TOP                          |
| 24   | DISTINCTROW + TOP                       |
| 48   | TOP *n* PERCENT                         |
| 50   | DISTINCT + TOP PERCENT                  |
| 56   | DISTINCTROW + TOP PERCENT               |

For `TOP n` and `TOP n PERCENT`, `Name1` carries the count.

### Attribute 5 — input tables (FROM)

One row per table, query, or UNION segment. `Name1` is the table or query
name; `Name2` is the alias (if any) or the UNION segment id.

### Attribute 6 — output columns (SELECT list)

| Flag    | Meaning                                     |
|---------|----------------------------------------------|
| 0       | Normal field or expression                   |
| 1       | Crosstab column heading field                |
| 2       | Crosstab row heading field                   |
| -32768  | Append VALUES literal                        |

The April 2026 addendum corrects the upstream docs: simple field references
go in `Expression` (not `Name1`); only the alias of a calculated column
goes in `Name1`.

### Attribute 7 — joins

| Flag | Join type           |
|------|---------------------|
| 1    | INNER JOIN          |
| 2    | LEFT (OUTER) JOIN   |
| 3    | RIGHT (OUTER) JOIN  |

`Name1` = left table, `Name2` = right table, `Expression` = the ON clause.

### Join reconstruction specification

Access stores a nested SQL join tree as ordered Attribute 5 input bindings plus
ordered Attribute 7 predicate fragments. Reconstruction must preserve the
tree implied by those ordered rows; treating Attribute 7 rows as unordered graph
edges is not sufficient.

**Bindings.** Attribute 5 rows define the available table references:

- `Name1` is the base table/query name.
- `Name2` is the alias when present; joins reference the alias, not the base
  name.
- For derived tables, `Expression` carries the inner `SELECT` and `Name2`
  carries the alias.

**Predicate fragments.** Attribute 7 rows are processed in `[Order]` sequence:

- `Flag` gives the join operator (`INNER`, `LEFT`, `RIGHT`).
- `Name1` and `Name2` are table references visible in the current join scope.
- `Expression` is one ON predicate fragment. Access can store one logical
  `ON (...) AND (...)` clause as multiple Attribute 7 rows.

**Logical join ownership.** A logical join node owns one or more Attribute 7
predicate fragments. A fragment belongs on the deepest join node whose left and
right operand scopes contain the two table references used by the fragment.
When a later Attribute 7 row would create a cycle in the already-built table
component, it is a secondary predicate for an existing logical join, not a new
join. Attach that secondary predicate to the primary join that introduced the
later-visible endpoint of the secondary predicate:

1. Track which primary join first introduced each table reference into the
   current component.
2. For a secondary predicate between `A` and `B`, compare the introduction
   order for `A` and `B`.
3. Attach the predicate to the join that introduced the later endpoint.
4. If neither endpoint has an introduction join (malformed/ambiguous storage),
   fall back to the deepest prior join that shares either endpoint and emit a
   warning if no valid owner is found.

The introduced-endpoint rule covers both nested-subtree directions. If a
secondary predicate references a table newly joined inside a subtree, it stays
with that nested join. If it references a root table newly joined to an existing
subtree, it stays with the root-to-subtree join. Attaching it to a later parent
join can produce Access SQL whose `ON` clause references a table outside the
joined operand; such SQL can fail with `JOIN expression not supported`.

**Validation invariants.**

- Every Attribute 7 `Name1`/`Name2` reference must resolve to an Attribute 5
  binding or alias.
- A generated join's `ON` clause must reference only tables visible in the
  join's left and right operands.
- Secondary predicates must be preserved; if no valid owner can be identified,
  validation should warn rather than silently moving or dropping the predicate.
- For query shapes with available dependencies, generated SQL must be accepted
  by Access and produce matching row count/schema against Access's
  `QueryDefs.SQL`.

### Attribute 8/9/10/11 — WHERE / GROUP BY / HAVING / ORDER BY

- **8** WHERE — `Expression` = the full criteria text (single row).
- **9** GROUP BY — one row per grouped field, `Name1` = field.
- **10** HAVING — `Expression` = the criteria.
- **11** ORDER BY — `Name1` = `D` or `d` for descending (blank = ascending);
  the field name lives in `Expression` per the April 2026 addendum
  (upstream docs say `Name2` — actual data uses `Expression`).

### Undocumented columns

- **`Order`** — 510-byte binary blob whose 4th byte (1-indexed) holds a
  1-based sequence number for rows within each Attribute group. Bytes 1–3
  are always `0x00`. The exporter must `ORDER BY Attribute, [Order]` to
  reproduce the canonical row order. Source: April 2026 addendum.
- **`MSysQueries.LvExtra`** — always NULL for query rows; not to be
  confused with `MSysObjects.LvExtra`, which is the design-layout blob.

### MSysObjects binary blobs (`LvProp`, `LvExtra`)

Both are `OLE Object` columns; only one (`LvProp`) is always present for
queries:

- **`LvProp`** — MR2 binary format (same as linked tables). Carries query
  properties (`ODBCTimeout`, `RecordsetType`, `DefaultView`, `Orientation`,
  `LogMessages`, etc.) and per-column metadata (`Name`, `AggregateType`,
  `ColumnWidth`, `ColumnHidden`). Parsed by `clsLvPropParser`. Always
  present. **`ReturnsRecords` is not stored in `LvProp`** for pass-through
  queries; export derives it from MSysQueries Attribute 1 Flag (`10` = no
  records) and writes it to the `.json` `QueryProperties` block when
  non-default.
- **`LvExtra`** — present only for queries last saved in Design View.
  Carries the design layout: window position, designer pane dimensions,
  table positions. Total size = 68 + (tableCount + 1) × 284 bytes. We
  serialize it into the `.json` companion's `DesignLayout` block and read
  it back on import. Header (68 bytes): magic `0x99 0x99 0xCE 0xAC`,
  12-byte `0xAA` padding, window RECT, state, designer pane RECT, grid
  origin, ColumnsShown, table count. Per-table entries (284 bytes each):
  five Longs (Left, Top, Right, Bottom, ScrollTop) + 132-byte
  null-terminated UTF-16LE name + 132-byte alias; one phantom
  all-zeros entry beyond the table count.
- **`Lv`** and **`LvModule`** — always NULL for queries; the parser
  ignores them.

### OptionFlag (`.json` companion)

`OptionFlag` is the JSON-serialized bitmask the new `.sql`/`.json`
pipeline writes; values match Attribute 3 above (e.g. 48 = TOP n PERCENT,
2 = DISTINCT, 1 = OutputAllFields). Stored in the `.json` companion so
the SQL string in the `.sql` file can stay clean (the importer combines
the SQL with `OptionFlag` when generating the `.qdef`).

## 3. Design View vs SQL View

### When Access uses each storage form

The MSysQueries data **differs** depending on the last-saved view, but the
difference is minimal — it does not encode the view itself. The view is
encoded in `MSysObjects.LvExtra`:

- **`MSysObjects.LvExtra IS NOT NULL`** ⇒ query was last saved in **Design
  View**. `Application.SaveAsText` emits a `.qdef` that begins with
  `Operation = N` (where N matches Attribute 1's Flag) and contains
  structured `InputTables` / `OutputColumns` / `Joins` blocks plus a
  layout block.
- **`MSysObjects.LvExtra IS NULL`** ⇒ query was last saved in **SQL
  View**. `Application.SaveAsText` emits a `.qdef` that begins with
  `dbMemo "SQL" = "..."` containing the raw SQL text and no structured
  blocks.

Differences in MSysQueries between the two saves (both forms still hold the
same query):

- SQL View saves omit Attribute 1 with Flag = 1 (SELECT) and Attribute 3
  with Flag = 0 (no options) — both are the default in SQL view, no need
  to store. Action queries (INSERT/UPDATE/DELETE/Make-Table) still keep
  Attribute 1 because Flag ≠ 1.
- If Attribute 3 Flag ≠ 0 (e.g., DISTINCT, TOP, PERCENT), the row is
  retained in SQL View because the Flag carries meaningful information.

Some shapes have no Design View representation and are SQL-View-only by
necessity:

- **UNION**, **Data Definition (DDL)**, **Pass-through** — Access itself
  refuses to display these in Design View.
- Non-equi joins (e.g., `ON A.x > B.y`) — display will fail.

### Our arbitration rule

`clsDbQuery.ImportNewFormat` decides which view to generate when building
the on-the-fly `.qdef` for `Application.LoadFromText`:

```vba
blnDesignView = cComposer.IsDesignerCompatible And _
    ((Not dLayout Is Nothing) Or cComposer.RequiresDesignView)
```

The rule has two clauses, both necessary:

1. **`IsDesignerCompatible`** (gate on the SQL itself). Returns `False`
   for shapes Access Design View cannot render at all (UNION, DDL,
   pass-through, non-equi joins). Importing an incompatible shape via
   Design View qdef yields `Resource failure` from `LoadFromText`.

2. **`(Not dLayout Is Nothing) Or RequiresDesignView`** (gate on intent +
   constraint). Either:
   - `dLayout` is present in the `.json` companion — meaning the original
     `MSysObjects.LvExtra` was non-NULL, meaning the query was originally
     saved in Design View, so we should preserve that, **or**
   - `RequiresDesignView` is true — meaning the SQL contains a shape
     `LoadFromText` cannot consume in SQL View qdef format (currently:
     multi-condition `ON` clauses, see § 6).

If neither clause holds, we emit a SQL View qdef. This is the safe default
because `LoadFromText` accepts SQL View qdefs for every shape that doesn't
fall into § 6's asymmetry list.

The current implementation lives at:

- `clsQueryComposer.IsDesignerCompatible` — line 532
- `clsQueryComposer.RequiresDesignView` — line 559 (with detailed
  procedure-header documenting the constraint)
- `clsQueryComposer.HasTopLevelBoolean` — line 2263 (helper that detects
  the multi-condition `ON` shape)
- `clsDbQuery.ImportNewFormat` — site of the `blnDesignView` decision

**Table-less scalar SELECT in Design View.** Queries such as `SELECT
Date() AS Today;` (no `FROM` clause) store Attribute 6 output-column
rows but no Attribute 5 input-table rows. When the `.json` companion
carries `DesignLayout`, the importer uses the Design View path and must:

1. Parse the field list even when no `FROM` is present (`ParseSelectQuery`
   calls `ParseFieldList` on the entire post-`SELECT` remainder — same as
   `ParseInsertQuery` already did for append queries without a source
   table).
2. Always emit an empty `Begin InputTables` / `End` block in the generated
   `.qdef`. Access `SaveAsText` does this for Design View queries;
   omitting the block causes `LoadFromText` to silently drop
   `OutputColumns`, producing `SELECT FROM ;` on the next export.

The SQL View path (no `DesignLayout`) stores the raw SQL in `dbMemo "SQL"`
and is unaffected. Regression fixtures:
`qryRegressionScalarNoTable` (SQL View) and
`qryRegressionScalarNoTableDesignView` (Design View).

## 4. What our parser handles

Each row below has a canonical fixture under `Testing/Fixtures/queries/`.
A fixture entry means the round-trip harness verifies the shape on every
test run; a regression entry additionally implies "if this regresses we
have re-introduced a known bug." See
[Testing/Fixtures/README.md](../Testing/Fixtures/README.md) for the harness
contract.

| Shape                                       | Fixture                                                                                                                  |
|---------------------------------------------|--------------------------------------------------------------------------------------------------------------------------|
| Plain SELECT                                | [select/qryCars.sql](../Testing/Fixtures/queries/select/qryCars.sql)                                                     |
| SELECT with WHERE                           | [select/qryCarsFiltered.sql](../Testing/Fixtures/queries/select/qryCarsFiltered.sql)                                     |
| Aggregate / GROUP BY / Count                | [select/qryCarColoursCount.sql](../Testing/Fixtures/queries/select/qryCarColoursCount.sql)                               |
| INNER JOIN                                  | [select/qryCurrencyExchangeINNER.sql](../Testing/Fixtures/queries/select/qryCurrencyExchangeINNER.sql)                   |
| INNER JOIN with WHERE                       | [select/qryCurrencyExchangeINNERFiltered.sql](../Testing/Fixtures/queries/select/qryCurrencyExchangeINNERFiltered.sql)   |
| LEFT JOIN                                   | [select/qryCurrencyExchangeLEFT.sql](../Testing/Fixtures/queries/select/qryCurrencyExchangeLEFT.sql)                     |
| Mixed LEFT/RIGHT outer joins                | [regression/qryRegressionMixedOuterJoin.sql](../Testing/Fixtures/queries/regression/qryRegressionMixedOuterJoin.sql)     |
| Self-join, fully aliased (`AS a`/`AS b`)    | [regression/qryCurrencyCrossRates.sql](../Testing/Fixtures/queries/regression/qryCurrencyCrossRates.sql)                 |
| Self-join, unaliased (`_1` synthetic alias) | [regression/qryRegressionSelfJoinUnaliased.sql](../Testing/Fixtures/queries/regression/qryRegressionSelfJoinUnaliased.sql) |
| 3-table inner-join chain                    | [regression/qryRegressionStrandedAlias.sql](../Testing/Fixtures/queries/regression/qryRegressionStrandedAlias.sql)       |
| Multi-condition `ON` (parenthesized AND/OR) | [regression/qryRegressionMultiCondJoin.sql](../Testing/Fixtures/queries/regression/qryRegressionMultiCondJoin.sql)       |
| Cross-subtree `ON` predicate placement      | [regression/qryRegressionJoinPredicatePlacement.sql](../Testing/Fixtures/queries/regression/qryRegressionJoinPredicatePlacement.sql) |
| Aggregate cross-subtree `ON` placement      | [regression/qryRegressionJoinPredicatePlacementAggregate.sql](../Testing/Fixtures/queries/regression/qryRegressionJoinPredicatePlacementAggregate.sql) |
| Cross-table multi-condition `ON`            | [regression/qryRegressionCrossTableOn.sql](../Testing/Fixtures/queries/regression/qryRegressionCrossTableOn.sql)                                 |
| Cross-table `ON` emitter table pairs       | [regression/qryRegressionCrossTableOnEmitter.sql](../Testing/Fixtures/queries/regression/qryRegressionCrossTableOnEmitter.sql)                   |
| Nested secondary `ON` predicate placement   | [regression/qryRegressionNestedSecondaryJoinPredicate.sql](../Testing/Fixtures/queries/regression/qryRegressionNestedSecondaryJoinPredicate.sql) |
| Disconnected join graph (Cartesian + chain) | [regression/qryRegressionDisconnectedComponents.sql](../Testing/Fixtures/queries/regression/qryRegressionDisconnectedComponents.sql) |
| `IN (SELECT ...)` subquery                  | [regression/qryRegressionFindDuplicates.sql](../Testing/Fixtures/queries/regression/qryRegressionFindDuplicates.sql)     |
| Derived table in `FROM` (`%$##@_Alias`)     | [regression/qryRegressionFromSubquery.sql](../Testing/Fixtures/queries/regression/qryRegressionFromSubquery.sql)         |
| Quoted identifiers / brackets               | [regression/qryRegressionQuotes.sql](../Testing/Fixtures/queries/regression/qryRegressionQuotes.sql)                     |
| Backslash literals in string concat         | [regression/qryRegressionBackslash.sql](../Testing/Fixtures/queries/regression/qryRegressionBackslash.sql)               |
| `TOP N PERCENT`                             | [regression/qryRegressionTopPercent.sql](../Testing/Fixtures/queries/regression/qryRegressionTopPercent.sql)             |
| Scalar no-table SELECT                      | [regression/qryRegressionScalarNoTable.sql](../Testing/Fixtures/queries/regression/qryRegressionScalarNoTable.sql)       |
| Scalar no-table SELECT (Design View)        | [regression/qryRegressionScalarNoTableDesignView.sql](../Testing/Fixtures/queries/regression/qryRegressionScalarNoTableDesignView.sql) |
| Explicit `table.*` plus all-fields `*`      | [regression/qryRegressionExplicitAndAllFields.sql](../Testing/Fixtures/queries/regression/qryRegressionExplicitAndAllFields.sql) |
| Make-Table (`SELECT ... INTO`)              | [regression/qryRegressionExternalMakeTable.sql](../Testing/Fixtures/queries/regression/qryRegressionExternalMakeTable.sql) |
| Query parameters (Attribute 2)               | [regression/qryRegressionParameterizedCrosstab.sql](../Testing/Fixtures/queries/regression/qryRegressionParameterizedCrosstab.sql) |
| INSERT INTO ... SELECT (Append)             | [append/qryAppendCars.sql](../Testing/Fixtures/queries/append/qryAppendCars.sql)                                         |
| Scalar append without source table          | [regression/qryRegressionScalarAppendNoTable.sql](../Testing/Fixtures/queries/regression/qryRegressionScalarAppendNoTable.sql) |
| UPDATE                                      | [update/qryUpdateCarsPrice.sql](../Testing/Fixtures/queries/update/qryUpdateCarsPrice.sql)                               |
| UPDATE DISTINCTROW                          | [regression/qryRegressionUpdateDistinctRow.sql](../Testing/Fixtures/queries/regression/qryRegressionUpdateDistinctRow.sql) |
| DELETE                                      | [delete/qryDeleteUnusedCurrencies.sql](../Testing/Fixtures/queries/delete/qryDeleteUnusedCurrencies.sql)                 |
| DELETE DISTINCTROW                          | [regression/qryRegressionDeleteDistinctRow.sql](../Testing/Fixtures/queries/regression/qryRegressionDeleteDistinctRow.sql) |
| TRANSFORM (Crosstab)                        | [crosstab/qryCarsCrosstab.sql](../Testing/Fixtures/queries/crosstab/qryCarsCrosstab.sql)                                 |
| UNION                                       | [union/qryUnionMakers.sql](../Testing/Fixtures/queries/union/qryUnionMakers.sql)                                         |
| UNION with global ORDER BY                  | [regression/qryRegressionUnionOrderBy.sql](../Testing/Fixtures/queries/regression/qryRegressionUnionOrderBy.sql)         |
| Data Definition (DDL)                       | [ddl/qryCreateTempTable.sql](../Testing/Fixtures/queries/ddl/qryCreateTempTable.sql)                                     |
| MSysQueries / MSysObjects introspection     | [select/qryMSysQueries.sql](../Testing/Fixtures/queries/select/qryMSysQueries.sql)                                       |

## 5. Known gaps — what we do not yet handle

Each row below is a query shape documented in the upstream Riddington
articles or `MSysQueriesExamples` repo that has **no fixture in this
repo**. Behaviour through our import/export pipeline is therefore
**unverified**; the parser may produce correct output, may silently drop
information, or may fail loudly. Treat each row as an opportunity to
contribute a fixture (see the bug-as-fixture workflow in
[Testing/Fixtures/README.md](../Testing/Fixtures/README.md)).

| Shape                                              | Status / what would be needed                                                                                                  |
|----------------------------------------------------|--------------------------------------------------------------------------------------------------------------------------------|
| Query parameters beyond Text (`Attribute 2`)       | Text parameters are pinned by `qryRegressionParameterizedCrosstab`. Riddington documents additional DAO data-type flags (Boolean, Byte, Integer, Long, Currency, Single, Double, Date/Time, OLE, Memo, GUID, BigInt); add fixtures if those are observed in the wild. |
| `WITH OWNERACCESS OPTION` (Attribute 3 Flag = 4)   | No fixture. Verify the option flag round-trips (currently unknown whether it's preserved in `OptionFlag`).                     |
| Multi-value field (MVF) references                 | No fixture. Riddington Part 1 § 21. Likely shows up in Attribute 12 Flag = 2.                                                  |
| Attachment field references                        | No fixture. Same Attribute 12 Flag = 2 family as MVF.                                                                          |
| Long-text version-history references               | No fixture. Attribute 12 Flag = 1. Likely out of scope (version history is an Access-managed feature).                         |
| TEMP queries (`MSysObjects.Flags = 3`)             | Intentionally not exported. Access auto-creates these (`~sq_*` names) for form/report record sources; they're regenerated on demand. The `clsDbQuery.GetAllFromDB` enumerator filters them out. |
| Deleted-query tombstones (`~TMPCLP*`)              | Intentionally not exported. Access marks recently-deleted queries this way; they're permanently removed on compact.            |
| Pass-through queries                               | Covered by `passthrough/qryPassThroughNoConnect`, `qryPassThroughReturnsRecords` (Attribute 1 Flag 8), and `qryPassThroughNoRecords` (Flag 10 / `ReturnsRecords=false`, issue #724). SQL is the verbatim Attribute 1 `Expression`; connect is Attribute 1 `Name1` (or Attribute 4 when present). |
| Scalar `SELECT 1 ... AS X` subqueries in projection | No fixture. Riddington Part 1 § 27 example uses `Exists (SELECT 1 ...)` inside a DELETE; only the outer DELETE shape is currently exercised. |
| `INSERT INTO ... VALUES (...)` (literal append)    | No fixture. Differs structurally from `INSERT INTO ... SELECT` (uses Attribute 6 Flag = -32768 for VALUES literals).           |
| Non-equi joins (`A.x > B.y`)                       | No fixture. Cannot be displayed in Design View — would be SQL-View-only and may need to land alongside the multi-cond `ON` asymmetry in § 6. |
| External-database joins (`IN '...'` clause)        | No fixture. Attribute 4 carries the connect string; verify it round-trips.                                                     |

When adding a fixture for any of the above, follow the contribution
workflow in [Testing/Fixtures/README.md § Bug-as-fixture](../Testing/Fixtures/README.md).
If the new fixture exposes a parser bug, fix it in `clsQueryComposer`
under the protection of the new fixture; if it exposes an
`Application.LoadFromText` quirk, document it in § 6 below.

## 6. Findings unique to our pipeline

Findings below were discovered through running the round-trip harness
and are **not** in the upstream Riddington documentation. They are
specific to this add-in's use of `Application.LoadFromText` and
`Application.SaveAsText` for object I/O.

### LoadFromText / SaveAsText asymmetry for multi-condition `ON`

**Symptom.** A query with a multi-condition `ON` clause —
`JOIN tblB ON (tblA.x = tblB.x) AND (tblA.y = tblB.y)` — that Access
itself displays in either view, when emitted as a SQL View qdef
(`dbMemo "SQL" = "..."`), is **rejected by `Application.LoadFromText`**
even though `Application.SaveAsText` will happily produce the same text.

**Empirical evidence.**

1. Created the query in Access via Design View; saved.
2. `Application.SaveAsText` produced a SQL View qdef
   (`dbMemo "SQL" = "..."`) — Access stored the query in SQL View.
3. Re-running `Application.LoadFromText` on that exact same file failed
   with `Error 2128 — Version Control System encountered errors while
   importing`.
4. Generating a Design View qdef from the same SQL (with structured
   `Joins` rows, one per ON sub-condition) loaded successfully.

**Implication.** The composer cannot rely on Access's own SQL View
representation. For multi-condition `ON` it must emit Design View qdef
format regardless of whether the source query had `LvExtra` populated.

**Implementation.** `clsQueryComposer.HasTopLevelBoolean` (line 2263)
detects `AND` / `OR` at the top level of an ON expression after stripping
outer parentheses. When it returns `True` for any join expression,
`m_blnRequiresDesignView` is set, exposed via the `RequiresDesignView`
property (line 559), and consumed by the arbitration rule in
`clsDbQuery.ImportNewFormat` (see § 3).

**Pinned by:** [regression/qryRegressionMultiCondJoin.sql](../Testing/Fixtures/queries/regression/qryRegressionMultiCondJoin.sql)
+ sibling `.notes.md`.

### `TOP N PERCENT` lives in SQL View in Access

**Symptom.** When the importer attempted to emit a Design View qdef for a
`TOP N PERCENT` query that lacked `DesignLayout` in its `.json`,
`Application.LoadFromText` returned `Resource failure`.

**Cause.** Access itself stores `TOP N PERCENT` queries in SQL View
(`MSysObjects.LvExtra IS NULL`, `Application.SaveAsText` emits
`dbMemo "SQL" = "..."`). The composer was previously generating a Design
View qdef whenever `IsDesignerCompatible` returned `True`, regardless of
whether the original was actually Design View. For `TOP N PERCENT`,
Access's own Design View qdef format apparently can't represent the
`PERCENT` flag in the way `LoadFromText` expects, hence the failure.

**Implication.** `IsDesignerCompatible` is necessary but not sufficient
for Design View import. Absence of `DesignLayout` in the `.json` is
authoritative for SQL View intent.

**Implementation.** The arbitration rule in § 3 — Design View only when
`(Not dLayout Is Nothing) Or RequiresDesignView` — is what fixes this.
`TOP N PERCENT` queries with no `DesignLayout` and no multi-cond `ON`
correctly fall through to SQL View, which `LoadFromText` accepts.

**Pinned by:** [regression/qryRegressionTopPercent.sql](../Testing/Fixtures/queries/regression/qryRegressionTopPercent.sql)
+ sibling `.notes.md`.

### Derived tables in `FROM` live in MSysQueries `Expression`, not `Name1`

**Symptom.** A query whose `FROM` clause contains a derived table
(subquery), e.g. `FROM (SELECT ... FROM ...) AS [%$##@_Alias]`, exported
through the v5 pipeline lost the entire subquery. The emitted `.sql`
collapsed to `FROM   AS % $ ##@_Alias;` (empty table reference, alias
also corrupted by the formatter).

**Cause.** Two compounding bugs in `clsQueryComposer.ReconstructSQL` /
`BuildFromClause` (line 657):

1. **Wrong field read.** Access stores derived tables in MSysQueries with
   `Attribute = 5`, `Name1 = NULL`, `Name2 = <alias>`,
   `Expression = <inner SELECT>`. The composer read `Name1` as the table
   name and ignored `Expression` (which is only consumed by the UNION
   branch). With `Name1` empty the FROM emitter produced no table
   reference at all.
2. **Alias bracketing too narrow.** `BracketIfNeeded` only bracketed
   names containing spaces or matching the reserved-word list. The
   `%$##@_Alias` placeholder has neither, so it passed through
   unbracketed and `clsSqlFormatter` then tokenized `%`, `$`, `#`, `@`
   as separate operators -- producing `% $ ##@_Alias` in the output.

**How Access generates the `%$##@_Alias` placeholder.** When the query
designer wraps an unnamed subquery in the FROM clause, Access auto-names
it `%$##@_Alias`. User-written SQL with a named subquery
(`FROM (SELECT ...) AS sub`) hits the same code path with `Name2 = "sub"`
instead. The placeholder is not the bug -- it's the most common shape
to trip it.

**Fix.** A new helper `FormatInputTableName` in `clsQueryComposer`
detects the empty-`Name1`/non-empty-`Expression` case and emits
`(<expression>)` so the FROM clause reads `(SELECT ...) AS [<alias>]`.
`BracketIfNeeded` was extended via `HasNonIdentChars` to bracket any
identifier containing characters outside `[A-Za-z0-9_]`, and to pass
already-parenthesized expressions (derived-table subqueries) through
unchanged.

**Why `IsDesignerCompatible` is still the right gate.** `HasSubqueries`
(line 511 of the post-fix file) returns True whenever any input table
name contains `(` or starts with `SELECT `, so the importer continues to
emit SQL View qdef (`dbMemo "SQL"`) for these. Access's `LoadFromText`
accepts SQL View qdefs containing FROM-clause subqueries (verified by
the harness); the legacy 4.x `.bas` shape that triggered the original
user report --
`InputTables.Name = "<entire SELECT>"` / `Alias = "%$##@_Alias"` -- is
specific to `Application.SaveAsText`'s Design View qdef format and is
the asymmetry the new pipeline was already designed to sidestep.

**Pinned by:** [regression/qryRegressionFromSubquery.sql](../Testing/Fixtures/queries/regression/qryRegressionFromSubquery.sql)
+ sibling `.notes.md`.

### Design View layout cannot be restored for non-designer-compatible queries

**Symptom.** After a full build from source, queries that were imported
via SQL View qdef (`dbMemo "SQL"`) — e.g. FROM-clause subqueries, UNION,
DDL, pass-through — have `MSysObjects.LvExtra IS NULL`. The user must
manually switch to Design View (for shapes that support it) to get
a layout; this layout is Access's auto-default, not the original.

**Why the original layout cannot be programmatically restored.**
Three approaches were tested (April 2026); all fail:

1. **Direct DAO write to `MSysObjects.LvExtra`** — `DAO.Recordset.Edit`
   returns error 3027 ("Cannot update. Database or object is read-only")
   even with `dbOpenDynaset`.
2. **ADO write after `GRANT UPDATE ON MSysObjects TO Admin`** — the
   `GRANT` succeeds but the subsequent `ADODB.Recordset.Update` returns
   error −2147217911 (same "read-only" message). The ACE engine
   hardcodes `MSysObjects` write protection at a level below the
   Jet security model.
3. **`DoCmd.OpenQuery acViewDesign` + `DoCmd.Close acSaveYes`** — this
   *does* populate `LvExtra` (verified: 636 bytes for a 2-table
   derived-table query), but only when Access dirties the query during
   the save. For queries whose output columns lack explicit `AS` aliases,
   Access auto-renames them to `Expr1, Expr2, …`, which breaks
   downstream consumers. The original SQL can be restored afterward via
   `qdef.SQL = originalSQL` (LvExtra survives the SQL setter), but the
   layout written is always Access's auto-default (one table box at
   position 72,18), not the original coordinates from the `.json`
   companion. For queries that already have explicit aliases, Access
   does not dirty the query, so LvExtra stays NULL.

**Why this is accepted as a known limitation.** The purpose of
preserving `DesignLayout` in the `.json` companion is to restore the
*user's* layout — table positions, window size, designer pane scroll.
Since none of the available write paths can place arbitrary coordinates
into `LvExtra`, auto-generating a default layout provides no value over
the user simply switching from SQL View to Design View manually (which
produces the same default layout in one click). Adding automated
complexity for a result the user can trivially achieve themselves is not
justified.

**What the user sees.** After rebuilding a database that contained a
FROM-subquery query originally saved in Design View:

- The query opens in SQL View (because `LvExtra` is NULL).
- The SQL is correct and fully functional.
- The user can switch to Design View at any time; Access generates a
  fresh default layout.
- Re-exporting captures the new layout in `DesignLayout`, which will
  differ from the original (source-control diff on first re-export,
  then stable).

### `ForceImportOriginalQuerySQL` is legacy-only

**Status.** The `ForceImportOriginalQuerySQL` option in `clsOptions`
(default `False`) is consulted **only** by
`clsDbQuery.ImportLegacyFormat` for legacy `.qdef` / `.bas` source
files. The new `.sql` / `.json` pipeline (`ImportNewFormat`) does
**not** read it.

**Why it was originally added.** Some rare query shapes can be exported
natively by `Application.SaveAsText` but cannot be re-imported by
`Application.LoadFromText`. The option lets the importer fall back to
setting `QueryDefs(name).SQL` directly from a sidecar `.sql` file,
losing visual designer state but recovering an unusable `.qdef`.

**Why it stays in place.** It remains a safety net for users on the
legacy path. **No deprecation plan.** When the same kind of
asymmetry was discovered in the new pipeline (multi-condition `ON`,
above), we addressed it structurally with `RequiresDesignView` rather
than by reusing the option, because the new pipeline can choose its
qdef shape rather than being handed a pre-baked one.

**Implementation.** `clsDbQuery.ImportLegacyFormat` (handles `.qdef` /
`.bas`); `clsOptions.ForceImportOriginalQuerySQL`. The user-facing
description of the option is at
[Wiki/Options.md](../Wiki/Options.md).

### Cross-table `ON` condition requires per-condition `LeftTable`/`RightTable` in Design View qdef

**Symptom.** A query with a compound `ON` clause whose individual
conditions reference different table pairs — e.g.
`ON (A.x = B.x) AND (A.x = C.x)` — imports successfully via
`Application.LoadFromText` but fails at execution time with DAO error
3082 ("JOIN operation refers to a field that is not in one of the joined
tables"). The error occurs when one of the referenced tables also
appears inside a saved subquery referenced in another condition of the
same compound `ON`.

**Root cause.** Access stores each compound `ON` condition as a
separate Attribute 7 row in `MSysQueries`, each with its own
`Name1`/`Name2` (the specific table pair for that condition). When the
VCS emitter generated a Design View `.qdef`, it reused the parent
join's `LeftTable`/`RightTable` for all split conditions. For the
second condition (`A.x = C.x`), the emitted `RightTable` was `B`
(the parent join's right table) instead of `C`.

`Application.LoadFromText` accepted the `.qdef` without error, but the
internal `MSysQueries` storage then had `Name2 = B` for a condition
referencing `C`. When Access compiled the query at execution time, it
looked for `C.x` within the scope of table `B` — and when `B` is a
saved query that itself references a table sharing columns with `C`,
the engine's scope resolution produced the ambiguity that triggers
error 3082.

**Empirical evidence.**

1. Exported a production query with `SaveAsText` — the native `.qdef`
   had per-condition `LeftTable`/`RightTable` (correct).
2. The VCS emitter produced a `.qdef` with the parent join's
   `RightTable` on all conditions (incorrect).
3. `LoadFromText` on the VCS-emitted `.qdef` succeeded (no error).
4. Executing the reimported query failed with error 3082.
5. Manually correcting the `RightTable` to match the condition's
   actual table pair, then re-importing, produced a query that
   executed correctly with the expected row count.

**Implication.** Unlike the multi-condition `ON` `LoadFromText`
asymmetry (above), this bug was in the VCS emitter, not in Access.
`LoadFromText` accepts the wrong `LeftTable`/`RightTable` silently —
the error manifests only at query execution, making it hard to
diagnose.

**Implementation.** `clsQueryComposer.EmitDesignViewQdef` now uses
`ExtractTableFromOnSide` to derive the correct `LeftTable`/`RightTable`
for each individual condition in a split compound `ON`, falling back to
the parent join's tables only if extraction fails.

**Pinned by:** [regression/qryRegressionCrossTableOn.sql](../Testing/Fixtures/queries/regression/qryRegressionCrossTableOn.sql)
\+ sibling `.notes.md`.

## 7. References

**External:**

- [Colin Riddington — Explaining Queries Part 1](https://www.isladogs.co.uk/explaining-queries/index.html)
- [Colin Riddington — Explaining Queries Part 2 (Design vs SQL view)](https://www.isladogs.co.uk/explaining-queries-2/index.html)
- `MSysQueriesExamples` companion repo (`C:\Repos\MSysQueriesExamples\`) —
  worked-example queries, live test database, April 2026 binary-format
  addendum.
- [Access Database Engine — Recover Deleted Database Objects](https://www.isladogs.co.uk/recover-deleted-objects/) (referenced from Riddington Part 1 § 29).

**In this repo:**

- [DECISIONS.md](../DECISIONS.md) — entries tagged with `clsQueryComposer`
  or `clsDbQuery` carry the rationale behind every choice this doc
  describes. Notable: `2026-04-14 — SQL reconstruction fidelity`,
  `2026-04-14 — Round-trip import with Design View / SQL View fallback`,
  `2026-04-14 — Column metadata and property serialization strategy`,
  `2026-04-20 — Wrap query composer pipeline in CatchAny error handling`,
  `2026-04-10 — Deterministic query export with performance optimization`.
- [Testing/Fixtures/README.md](../Testing/Fixtures/README.md) — the
  round-trip harness contract and the bug-as-fixture contribution workflow.
- [Testing/Fixtures/queries/](../Testing/Fixtures/queries/) — the fixture
  corpus referenced from § 4. Each `regression/*.notes.md` carries
  shape-specific context for the fixture next to it.
- [Version Control.accda.src/AGENTS.md § Query Files](../Version%20Control.accda.src/AGENTS.md) —
  user-facing description of the `.sql` / `.json` source format and the
  "Before changing the query parser" pointer.
- [Version Control.accda.src/modules/Utility/clsQueryComposer.cls](../Version%20Control.accda.src/modules/Utility/clsQueryComposer.cls)
  — the composer; key entry points are `DecomposeSQL`, `GenerateQdef`,
  `IsDesignerCompatible`, `RequiresDesignView`, `HasTopLevelBoolean`.
- [Version Control.accda.src/modules/Components/clsDbQuery.cls](../Version%20Control.accda.src/modules/Components/clsDbQuery.cls)
  — the component class; key entry points are `ExportNewFormat`,
  `ImportNewFormat`, `ImportLegacyFormat`.
- [Version Control.accda.src/modules/Utility/clsLvExtraParser.cls](../Version%20Control.accda.src/modules/Utility/clsLvExtraParser.cls)
  — the binary parser for the `MSysObjects.LvExtra` design-layout blob.
