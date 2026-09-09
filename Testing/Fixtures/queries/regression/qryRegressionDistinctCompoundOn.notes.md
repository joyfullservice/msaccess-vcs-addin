# qryRegressionDistinctCompoundOn

**Pins:** `SELECT DISTINCT` plus a compound `ON` clause must keep the
same semantics on both import views.

Compound `ON` forces Design View (`RequiresDesignView`) even when the
companion JSON has no `DesignLayout`. SQL is the source of truth for
DISTINCT; `OptionFlag: 2` agrees with the `.sql` so a stale JSON bit
cannot add or remove the modifier. A disagreeing pair is covered by
`clsTestQueryComposer` unit tests, not this fixture.
