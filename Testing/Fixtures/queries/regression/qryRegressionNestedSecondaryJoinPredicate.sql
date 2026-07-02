SELECT
  tblCars.ID,
  tblCarsModel.Model,
  tblCarsColour.Colour,
  tblVehicles.VehicleType
FROM
  (
    tblCars
    INNER JOIN (
      (
        tblCarsModel
        INNER JOIN tblVehicles ON tblCarsModel.ID = tblVehicles.ID
      )
      INNER JOIN tblCarsColour ON (
        tblVehicles.ID = tblCarsColour.ID
      )
      AND (
        tblCarsModel.ModelID = tblCarsColour.ID
      )
    ) ON tblCars.ID = tblCarsModel.ID
  )
  INNER JOIN tblCarsPrice ON tblCars.ID = tblCarsPrice.ID;
