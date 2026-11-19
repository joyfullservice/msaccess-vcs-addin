# qryRegressionReservedAliasDesignView

**Pins:** from export format 5.2.0, the SQL builder brackets an output alias
exactly when Access's query designer does.

Access's designer brackets a wider set of words than the published Jet
reserved-word list: `Option`, `First`, `Second`, and `To` among them.
`AS Option` unbracketed is rejected outright (error 3141), so an export that
drops those brackets does not paste back into SQL View. The designer does
**not** bracket `Name`, which shows the two lists differ in both directions.

`MSysQueries` stores every alias without brackets, whether the query was saved
from SQL View or the designer, so the builder decides the brackets either way.
This fixture carries a `DesignLayout`, so it also imports through the
generated `.qdef`; the SQL View companion `qryRegressionReservedAlias` covers
the same aliases from saved SQL.

When this fixture starts failing, the regression is likely either:

- A word dropped from, or added to, the designer list in `clsSqlSyntax`.
- The alias path calling `BracketIfNeeded` without the designer-word flag.
