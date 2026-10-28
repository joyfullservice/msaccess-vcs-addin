# qryRegressionGridOrigin

Regression canary for asymmetric nonzero query-design grid origin coordinates.

## Shape

Design layout JSON with `GridLeft: 0` and `GridTop: 720` (not the all-zeros default).

## What broke

`clsLvExtraParser` read the grid-origin pair from the `LvExtra` blob in blob order and
labelled it `(gridLeft, gridTop)`. The blob actually stores **Top first** — the reverse
of every other RECT in the format. The two values therefore swapped on every round trip.

The qdef layout block runs the other way: `LoadFromText` requires `Left` before `Top`
and rejects the entire Design View import with `Expected: 'Left'. Found: Top.` if the
emitter reverses them. That failure is silent in the sense that the importer falls back
to SQL View, so the query still imports but loses its whole layout.

## What must stay true

- The blob reader assigns the first Long to `gridTop` and the second to `gridLeft`.
- The qdef emitter writes `Left` before `Top`.
- A second export preserves `GridLeft` and `GridTop` without swapping.
