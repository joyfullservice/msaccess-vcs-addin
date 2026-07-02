# qryRegressionNestedSecondaryJoinPredicate

**Pins:** secondary predicates inside a nested subtree must merge into the join
that introduced the nested endpoint, not the later parent join that happens to
share the other endpoint.

A production validation run exposed this when a predicate between two tables
inside a nested subtree was promoted to the parent root-to-subtree join. The
generated SQL was structurally different and can be rejected by ODBC.

This fixture specifically protects the introduced-endpoint ownership rule: when
a secondary predicate is detected, the composer attaches it to the join that
introduced the later-visible endpoint in the current join component.
