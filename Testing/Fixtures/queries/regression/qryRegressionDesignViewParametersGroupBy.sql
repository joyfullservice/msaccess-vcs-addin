PARAMETERS [Min Qty] Long;
SELECT
  tblParamSample.Category,
  Count(tblParamSample.ID) AS Cnt
FROM
  tblParamSample
WHERE
  (
    (
      (tblParamSample.Qty) >= [Min Qty]
    )
  )
GROUP BY
  tblParamSample.Category
ORDER BY
  tblParamSample.Category;
