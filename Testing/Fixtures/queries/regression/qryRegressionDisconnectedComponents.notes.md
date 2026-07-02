# qryRegressionDisconnectedComponents

**Pins:** legitimate multi-component join graphs (one connected join
chain plus an unrelated Cartesian table) must round-trip cleanly
without false "out of order" or "stranded" warnings.

Cartesian products of unrelated lookup tables are uncommon but valid;
the harness must not falsely treat them as bugs. The current
`BuildFromClause` multi-pass attachment correctly emits these as
intentional Cartesian appends, and this fixture asserts that the SQL
produced is byte-stable across two passes.

When this fixture starts failing, the regression is likely either:

- The "stranded join" warning logic falsely alarming on an
  intentional Cartesian, or
- The multi-pass attachment rewriting the intentional Cartesian as
  something else (e.g., dropping the second component entirely, or
  fabricating a spurious join condition between the two components).
