# qryRegressionJoinPredicatePlacementAggregate

**Pins:** aggregate queries with cross-subtree join predicates must keep the
secondary predicate on the join that introduced the referenced nested table.

This is the compact equivalent of a production aggregate-query failure where
the builder moved a predicate from the aggregate source join onto a later
lookup join.
