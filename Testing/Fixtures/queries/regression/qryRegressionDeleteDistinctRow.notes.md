# qryRegressionDeleteDistinctRow

**Pins:** DELETE action queries with Attribute 3 bit 8 must reconstruct as
`DELETE DISTINCTROW`, not plain `DELETE`.

A production validation run found several multi-table deletes where Access
retained `DISTINCTROW` but the builder emitted `DELETE`. Even when duplicate
handling is usually benign, dropping the option is not a faithful round-trip.
