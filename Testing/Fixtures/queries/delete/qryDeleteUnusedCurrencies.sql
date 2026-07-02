DELETE tblCurrencies.*,
tblCurrencyExchange.Currency
FROM
  tblCurrencies
  LEFT JOIN tblCurrencyExchange ON tblCurrencies.CurrencyCode = tblCurrencyExchange.Currency
WHERE
  (
    (
      (tblCurrencyExchange.Currency) Is Null
    )
  );
