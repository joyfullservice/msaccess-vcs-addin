# qryRegressionSingleQuotedLiteralSpacing

**Pins:** [issue #785](https://github.com/joyfullservice/msaccess-vcs-addin/issues/785) —
Design View import must preserve whitespace inside single-quoted SQL literals.

The structured `.qdef` path compacts output-column expressions to match Access's
native representation. That compaction must apply only to SQL outside literals:
spaces around `-` and `=` are data here, not formatter whitespace. The escaped
single quote and embedded double-quoted text also pin independent delimiter
tracking.

The companion carries `DesignLayout`, ensuring the round-trip harness exercises
the affected Design View path rather than preserving the raw SQL through a SQL
View memo.
