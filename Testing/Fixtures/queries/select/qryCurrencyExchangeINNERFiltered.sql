SELECT
  tblCurrencies.CurrencyCode,
  tblCurrencies.Currency,
  tblCurrencyExchange.Base,
  tblCurrencyExchange.Date,
  tblCurrencyExchange.Currency,
  tblCurrencyExchange.Rate
FROM
  tblCurrencies
  INNER JOIN tblCurrencyExchange ON tblCurrencies.CurrencyCode = tblCurrencyExchange.Currency
WHERE
  (
    (
      (tblCurrencyExchange.Base) = "GBP"
    )
    AND (
      (tblCurrencyExchange.Date) = #12/14/2018#
    )
  )
ORDER BY
  tblCurrencies.CurrencyCode;
