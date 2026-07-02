# qryRegressionSelfJoinUnaliased

**Pins:** self-join where one side keeps the original table name and
the other side uses Access's auto-generated `_1` alias suffix.

The existing `qryCurrencyCrossRates` fixture covers self-joins in
which both sides are explicitly aliased (`AS a` / `AS b`). This
fixture covers the more common shape Access generates when the user
drags the same table twice into the designer: one input keeps the
table name and the second gets `<table>_1`.

The MSysQueries representation differs subtly between the two shapes
(both-aliased writes two `Alias =` lines; one-aliased writes only
one). The composer must continue to handle both shapes without
conflating them or collapsing `Name1`/`Name2` to the same value.

This shape was responsible for the residual `left=X right=X` warnings
that survived the first round of self-join fixes.
