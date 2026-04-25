SELECT
  tblCars.Manufacturer,
  tblCarsModel.Model,
  tblCurrencies.CurrencyCode
FROM
  (
    tblCars
    INNER JOIN tblCarsModel ON tblCars.ID = tblCarsModel.ID
  ),
  tblCurrencies;
