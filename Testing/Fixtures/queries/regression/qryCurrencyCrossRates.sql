SELECT
  a.Date,
  a.Base,
  a.Currency AS Currency1,
  a.Rate AS Rate1,
  b.Currency AS Currency2,
  b.Rate AS Rate2,
  (b.Rate / a.Rate) AS CrossRate
FROM
  tblCurrencyExchange AS a
  INNER JOIN tblCurrencyExchange AS b ON (a.Base = b.Base)
  AND (a.[Date] = b.[Date])
WHERE
  (
    (
      (a.Currency)< [b].[Currency]
    )
  );
