SELECT
  tblCars.Manufacturer,
  tblCarsModel.Model,
  tblCarsColour.Colour,
  tblCarsPrice.Price
FROM
  (
    (
      (
        tblCars
        LEFT JOIN tblCarsModel ON (tblCars.ID = tblCarsModel.ID)
      )
      LEFT JOIN tblCarsColour ON (tblCars.ID = tblCarsColour.ID)
      AND (tblCarsColour.ID > 0)
    )
    LEFT JOIN tblCarsPrice ON (tblCars.ID = tblCarsPrice.ID)
    AND (
      tblCarsModel.ModelID = tblCarsPrice.ID
    )
  )
  LEFT JOIN tblVehicles ON tblCars.ID = tblVehicles.ID;
