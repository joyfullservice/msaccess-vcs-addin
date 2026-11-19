# qryRegressionReservedAlias

**Pins:** output aliases that are Access reserved/contextual keywords must be
bracketed when reconstructed from MSysQueries.

The SQL builder previously emitted `AS Key` for an alias stored as `Key`.
Access renders the same alias as `AS [Key]`, and the unbracketed form can be
parsed as a keyword in downstream formatting/import paths.

This fixture also covers aliases seen in production validation diffs
(`Action`, `Year`, `Names`, `Number`, and `Currency`) so the
reserved/contextual alias list does not drift back to the narrower original
set.

`MSysQueries` stores aliases without brackets even for a query saved from SQL
View, so this fixture exercises the builder's alias list, not the text as
written. The words added to that list at export format 5.2.0 (`Option`,
`First`, `Second`, `To`) and the unbracketed `Name` are pinned by the Design
View companion `qryRegressionReservedAliasDesignView`.
