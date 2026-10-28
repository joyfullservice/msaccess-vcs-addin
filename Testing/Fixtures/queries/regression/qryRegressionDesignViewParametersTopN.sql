PARAMETERS [Min Price] Currency;
SELECT TOP 5
  tblParamSample.ID,
  tblParamSample.Category,
  tblParamSample.Price
FROM
  tblParamSample
WHERE
  (
    (
      (tblParamSample.Price) >= [Min Price]
    )
  )
ORDER BY
  tblParamSample.Price;
