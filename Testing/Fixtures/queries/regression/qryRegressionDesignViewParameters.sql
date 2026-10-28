PARAMETERS [Enter ID] Long,
[Enter Category] Text (255);
SELECT
  tblParamSample.ID,
  tblParamSample.Category
FROM
  tblParamSample
WHERE
  (
    (
      (tblParamSample.ID) = [Enter ID]
    )
    AND (
      (tblParamSample.Category) = [Enter Category]
    )
  );
