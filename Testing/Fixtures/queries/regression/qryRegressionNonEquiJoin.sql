SELECT
  A.MeterFK,
  A.ReadingDate,
  A.Reading,
  B.ReadingDate,
  B.Reading,
  [a].[readingdate] - [b].[readingdate] AS days,
  [A].[reading] - [b].[reading] -(
    100000 *([a].[readingdate] = #11/12/2020#)
  ) AS used,
  [used] / [days] AS perday
FROM
  tblMeterReadings AS A
  INNER JOIN tblMeterReadings AS B ON (A.MeterFK = B.MeterFK)
  AND (A.Reading > B.ReadingDate)
WHERE
  (
    (
      (A.MeterFK) <= 1
    )
    AND (
      (B.Estimate) = False
    )
    AND (
      (A.Estimate) = False
    )
  )
ORDER BY
  A.MeterFK,
  A.ReadingDate;
