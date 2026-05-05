# qryRegressionAppendValues

**Pins:** VALUES-style INSERT INTO with literal SELECT expressions and
no source table.

Earlier versions pinned the Access `FROM ;` rendering for this shape. The
newer exporter preserves the effective no-source append form without adding
`FROM ;`, matching scalar append behavior found in production validation.

The important invariant is that the INSERT target column list and SELECT
literal expressions survive round-trip; the empty source clause must not cause
the field list to be skipped or corrupted.
