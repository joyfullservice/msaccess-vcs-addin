INSERT INTO tblSettings (ID, ItemName, ItemValue)
SELECT
  12 AS Expr1,
  "GeneratedAt" AS Expr2,
  CStr(
    Date()
  ) AS Expr3;
