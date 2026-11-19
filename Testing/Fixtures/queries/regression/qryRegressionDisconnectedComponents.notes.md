# qryRegressionDisconnectedComponents

**Pins:** legitimate multi-component join graphs (one connected join
chain plus an unrelated Cartesian table) must round-trip cleanly
without false "out of order" or "stranded" warnings.

Cartesian products of unrelated lookup tables are uncommon but valid;
the harness must not falsely treat them as bugs. `BuildJoinForestSql`
emits these as intentional Cartesian tables, and this fixture asserts
that the SQL produced is byte-stable across two passes.

From export format 5.2.0 the unjoined table comes first and the join
group follows without outer parentheses (`FROM X, A INNER JOIN B ON ...`),
which is how Access writes the statement. Earlier formats append it after
the parenthesized join group. The import parser reads both orders. A
nested join with an unjoined table, saved from the designer, is pinned by
`qryRegressionDisconnectedNestedJoinDesignView`.

When this fixture starts failing, the regression is likely either:

- The "stranded join" warning logic falsely alarming on an
  intentional Cartesian, or
- The join forest rewriting the intentional Cartesian as
  something else (e.g., dropping the second component entirely, or
  fabricating a spurious join condition between the two components).
