PARAMETERS pActive Bit,
[Start Date] DateTime,
[Min Price] Currency,
[Max Weight] Double;
SELECT
  tblParamTypes.ID,
  tblParamTypes.Label
FROM
  tblParamTypes
WHERE
  (
    (
      (
        tblParamTypes.IsActive = pActive
        AND tblParamTypes.OrderDate >= [Start Date]
        AND tblParamTypes.Price >= [Min Price]
        AND tblParamTypes.Weight <= [Max Weight]
      )
    )
  );
