# qryRegressionScalarAppendNoTable

**Pins:** scalar `INSERT INTO ... SELECT expr, expr` queries with no input table
must not gain `FROM ;` during reconstruction.

A production validation run exposed this on scalar append queries with
Attribute 6 output rows but no Attribute 5 input table rows. They are distinct
from older literal append shapes that Access can store with an empty `FROM ;`
clause.
