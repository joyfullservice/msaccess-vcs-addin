SELECT
  tblCurrencies.CurrencyCode,
  tblCurrencies.Currency,
  tblCurrencyExchange.Base,
  tblCurrencyExchange.Date,
  tblCurrencyExchange.Currency,
  tblCurrencyExchange.Rate
FROM
  tblCurrencies
  LEFT JOIN tblCurrencyExchange ON tblCurrencies.CurrencyCode = tblCurrencyExchange.Currency
WHERE
  (
    (
      (tblCurrencyExchange.Base)= "GBP"
    )
  )
ORDER BY
  tblCurrencies.CurrencyCode,
  tblCurrencyExchange.Date DESC;
