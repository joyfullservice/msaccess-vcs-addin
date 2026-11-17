SELECT
  tblCars.Manufacturer,
  tblCurrencies.CurrencyCode,
  Count(tblCarsColour.Colour) AS CountOfColour
FROM
  tblCurrencies,
  tblCars
  INNER JOIN (
    tblCarsModel
    INNER JOIN tblCarsColour ON tblCarsModel.ID = tblCarsColour.ID
  ) ON tblCars.ID = tblCarsModel.ID
WHERE
  (tblCars.Year = 2011)
  AND (tblCurrencies.CurrencyCode = "USD")
GROUP BY
  tblCars.Manufacturer,
  tblCurrencies.CurrencyCode
HAVING
  (Count(tblCarsColour.Colour)) > 1;
