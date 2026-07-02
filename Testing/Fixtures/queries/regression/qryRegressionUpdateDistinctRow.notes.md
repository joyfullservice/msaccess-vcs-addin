# qryRegressionUpdateDistinctRow

**Pins:** UPDATE action queries with Attribute 3 bit 8 must reconstruct as
`UPDATE DISTINCTROW`, not plain `UPDATE`.

This complements the DELETE fixture because the SELECT modifier path is not
used by action queries; the builder has to emit the option directly after the
action keyword.
