# qryRegressionCrossSubtreeMultiCondJoin

Regression canary for a multi-condition ON clause whose secondary predicate references
a table outside the immediate join pair (cross-subtree secondary predicate).

## Shape

Two subtrees joined so that the final join's ON clause carries two predicates: one between
the tables actually being joined, and one referencing a table from the opposite subtree.

## Canonicalization note

The baseline here is what Access produces, not what was hand-authored. Given a
multi-condition ON where one predicate spans subtrees, Access keeps the predicate belonging
to the join pair inside the ON clause and re-emits the cross-subtree predicate as a trailing
top-level `AND (...)` after the join. The stored baseline reflects that normalized form.

This fixture therefore asserts two things: that the emitter reproduces Access's normalized
shape rather than fighting it, and that a second round trip is a fixed point. It does not
assert preservation of the original hand-written predicate placement.

## Distinction

- `qryRegressionMultiCondJoin` — stacked LEFT joins with local multi-condition ON.
- `qryRegressionNestedSecondaryJoinPredicate` — right-nested chain with a cross-table ON.
- This fixture — inner subtree whose second ON predicate references a table joined later.
