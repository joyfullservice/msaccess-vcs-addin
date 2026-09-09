# qryRegressionUnionWordInLiteral

**Pins:** the word `UNION` inside a string literal must not be read as a set
operator.

UNION detection scanned for the keyword while tracking only double quotes and
parenthesis depth, so a criterion such as `Manufacturer = "UNION ALL"` was
enough to classify a plain `SELECT DISTINCT` as a union query. That marks the
query designer-incompatible and replaces the DISTINCT bit with the UNION ALL
flag, silently rewriting `OptionFlag` on export.

`OptionFlag: 2` is the DISTINCT the `.sql` actually carries. A misread would
export flag 1 instead, which the round-trip invariant now rejects in both
directions.
