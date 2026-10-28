PARAMETERS [Select Category] Text (255);
TRANSFORM
  Count(tblParamDetail.ID) AS Cnt
SELECT
  tblParamSample.Category
FROM
  tblParamSample
  INNER JOIN tblParamDetail ON tblParamSample.ID = tblParamDetail.SampleID
WHERE
  (
    (
      (tblParamSample.Category) = [Select Category]
    )
  )
GROUP BY
  tblParamSample.Category
PIVOT
  tblParamDetail.StatusID In (1, 2, 3);
