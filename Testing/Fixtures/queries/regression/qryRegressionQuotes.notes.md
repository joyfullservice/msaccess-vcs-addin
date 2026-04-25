# qryRegressionQuotes

**Pins:** simple multi-column SELECT with quoted identifiers in the .json
metadata must round-trip without spurious whitespace or quoting changes.

Trivial-looking SQL, but the companion `.json` carries column metadata that
historically diffed across rebuilds when whitespace/escaping rules in the
formatter or JSON serializer changed. Use this as the canonical "no
surprises" baseline for the SELECT path.
