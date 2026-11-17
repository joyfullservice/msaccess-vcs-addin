UPDATE
  tblCurrencies,
  tblCars
  INNER JOIN tblCarsPrice ON tblCars.ID = tblCarsPrice.ID
SET
  tblCarsPrice.Price = [tblCarsPrice].[Price] * [tblCurrencies].[Rate]
WHERE
  (tblCars.Year = 2011)
  AND (tblCurrencies.CurrencyCode = "EUR");
