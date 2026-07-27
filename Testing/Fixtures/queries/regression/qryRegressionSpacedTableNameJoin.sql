SELECT
  tblA.fldA,
  tblB.fldB
FROM
  [Car Models]
  INNER JOIN [Car Colours] ON [Car Models].fldA = [Car Colours].fldB;
