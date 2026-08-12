PARAMETERS [Enter ID] Long,
[New Price] Currency;
UPDATE
  tblParamSample
SET
  tblParamSample.Price = [New Price]
WHERE
  (
    (
      (tblParamSample.ID) = [Enter ID]
    )
  );
