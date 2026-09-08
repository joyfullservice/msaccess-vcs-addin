SELECT
  a.SampleDate,
  a.GroupCode,
  a.ItemCode AS Item1,
  a.ItemValue AS Value1,
  b.ItemCode AS Item2,
  b.ItemValue AS Value2,
  (b.ItemValue / a.ItemValue) AS ValueRatio
FROM
  tblMeasurements AS a
  INNER JOIN tblMeasurements AS b ON (a.GroupCode = b.GroupCode)
  AND (a.[SampleDate] = b.[SampleDate])
WHERE
  (
    (
      (a.ItemCode) < [b].[ItemCode]
    )
  );
