# qryRegressionScalarNoTableDesignView

**Pins:** table-less scalar SELECT queries saved in Design View must keep
their output columns through import and must emit an empty
`Begin InputTables` / `End` block in the generated `.qdef`.

Access stores this shape with Attribute 6 output-column rows and no
Attribute 5 input-table rows. The companion `.json` carries a
`DesignLayout` (non-null `MSysObjects.LvExtra`), so the importer
chooses the Design View path rather than SQL View.

Reverting either of these behaviors breaks the round trip:

- `ParseSelectQuery` must call `ParseFieldList` when the SQL has no
  `FROM` clause (not only when `FROM` is present or bare at end).
- `EmitDesignViewQdef` must always write the InputTables block, even
  when empty; omitting it causes `LoadFromText` to silently drop
  OutputColumns.

Complements `qryRegressionScalarNoTable`, which covers the same SQL
shape on the SQL View path (no `DesignLayout` in `.json`).

The `.qdef` baseline matches the emitter output validated by
`qdef_vs_fixture` (layout block omitted there by design; layout is
validated via `json_vs_fixture` after import/export).
