SELECT
  tblItems.Name AS ItemName,
  tblItemTypes.SubType AS ItemType,
  Func1([Name]) AS ComputedA,
  Func2([Name]) AS ComputedB
FROM
  tblItems
  INNER JOIN tblItemTypes ON (
    tblItems.Flags = tblItemTypes.Flags
  )
  AND (
    tblItems.Type = tblItemTypes.Type
  )
WHERE
  (
    (
      (tblItems.Flags) <> 3
    )
    AND (
      (tblItems.Type) = 5
    )
  )
ORDER BY
  tblItems.Name;
