SELECT
  [%$##@_Alias].Manufacturer,
  [%$##@_Alias].ModelLabel,
  [%$##@_Alias].Year,
  [%$##@_Alias].Price,
  UCase([%$##@_Alias].Colour) AS ColourUpper
FROM
  (
    SELECT
      tblCars.ID,
      tblCars.Manufacturer,
      tblCars.Year,
      (
        tblCars.Manufacturer & ' ' & tblCarsModel.Model
      ) AS ModelLabel,
      tblCarsColour.Colour,
      tblCarsPrice.Price
    FROM
      (
        (
          tblCars
          INNER JOIN tblCarsModel ON tblCars.ID = tblCarsModel.ID
        )
        INNER JOIN tblCarsColour ON tblCarsModel.ID = tblCarsColour.ID
      )
      INNER JOIN tblCarsPrice ON tblCarsColour.ID = tblCarsPrice.ID
    WHERE
      tblCars.Year >= 2000
      AND tblCarsPrice.Price > 0
      AND tblCars.ID IN (
        SELECT
          ID
        FROM
          tblCarsColour
        WHERE
          Colour IN ('Red', 'Blue')
      )
    ORDER BY
      tblCars.Manufacturer
  ) AS [%$##@_Alias];
