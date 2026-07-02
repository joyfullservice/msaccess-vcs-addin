# qryRegressionOwnerAccess

**Pins:** `WITH OWNERACCESS OPTION` must survive query round-trip via
Attribute 3 option bit 4.

The SQL text carries the option at the end of the statement, while the JSON
`OptionFlag` preserves the MSysQueries option bit for qdef generation.
