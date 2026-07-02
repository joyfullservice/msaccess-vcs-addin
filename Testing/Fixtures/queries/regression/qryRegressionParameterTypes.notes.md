# qryRegressionParameterTypes

**Pins:** Attribute 2 parameter rows with a non-text DAO data type must survive
round-trip.

This fixture covers Date/Time parameter declarations. Existing fixtures already
cover Text and Short. Additional DAO parameter types should be added as separate
fixtures after confirming Access accepts the exact declaration syntax through
`LoadFromText`.
