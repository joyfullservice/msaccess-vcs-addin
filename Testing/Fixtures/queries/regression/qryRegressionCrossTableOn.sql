SELECT
  tblCars.Manufacturer,
  tblCarsModel.Model,
  tblCarsColour.Colour,
  tblCarsPrice.Price,
  tblVehicles.VehicleType
FROM
  (
    tblCars
    INNER JOIN (
      (
        tblCarsModel
        INNER JOIN tblCarsColour ON tblCarsModel.ID = tblCarsColour.ID
      )
      INNER JOIN tblCarsPrice ON tblCarsModel.ModelID = tblCarsPrice.ID
    ) ON (tblCars.ID = tblCarsModel.ID)
    AND (
      tblCars.Manufacturer = tblCarsColour.Colour
    )
  )
  INNER JOIN tblVehicles ON tblCars.ID = tblVehicles.ID;
