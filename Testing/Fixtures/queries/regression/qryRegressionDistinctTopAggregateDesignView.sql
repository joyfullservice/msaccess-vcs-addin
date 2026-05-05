SELECT
  DISTINCT TOP 3 Max(
    Year([Date])-1
  ) AS PriorYear
FROM
  tblCurrencyExchange
ORDER BY
  Max(
    Year([Date])-1
  ) DESC;
