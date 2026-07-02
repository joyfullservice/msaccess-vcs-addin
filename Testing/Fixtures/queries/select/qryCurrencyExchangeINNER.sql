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
ORDER BY
  tblCurrencies.CurrencyCode,
  tblCurrencyExchange.Date DESC;
