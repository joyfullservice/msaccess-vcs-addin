# qryRegressionOwnerAccessLiteral

**Pins:** a real `WITH OWNERACCESS OPTION` clause survives a round trip even
when the same phrase appears as data in a criterion.

The clause is recognized by a syntax-aware scan, so the copy inside the string
literal is skipped and only the trailing clause sets `OptionFlag` bit 4. The
same scan is what stops `ApplyOptionFlagToSql` from appending a second clause
at export format 5.1.0 -- a plain `InStr` would find the literal, conclude the
clause was already present, and drop the modifier from the emitted `.sql`.
