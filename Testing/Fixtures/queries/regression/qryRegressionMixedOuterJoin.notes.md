# qryRegressionMixedOuterJoin

**Pins:** mixed LEFT + RIGHT outer joins must preserve their direction
through round-trip, even when the multi-pass attachment in
`BuildFromClause` has to invert subtree placement to attach a join
whose `rightTable` is already in the accumulator.

The current attachment loop supports an "inverted" branch: if the new
join's `rightTable` is already in the accumulator and its `leftTable`
is new, the new table is attached on the *left* side of the
accumulated subtree, with the join direction flipped (LEFT <-> RIGHT)
to keep outer-join semantics correct.

This fixture exercises both directions in a single query so a
regression that accidentally collapses RIGHT JOIN into LEFT JOIN
(or vice versa) during inverted attachment would show up immediately
as a Pass 2 != Pass 1 diff, instead of being noticed only when query
results change at runtime.
