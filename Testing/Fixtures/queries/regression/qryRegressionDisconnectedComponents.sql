SELECT
  tblCars.Manufacturer,
  tblCarsModel.Model,
  tblCurrencies.CurrencyCode
FROM
  tblCurrencies,
  tblCars
  INNER JOIN tblCarsModel
    ON tblCars.ID = tblCarsModel.ID;
