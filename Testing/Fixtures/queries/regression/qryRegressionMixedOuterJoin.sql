SELECT
  tblCars.Manufacturer,
  tblCarsModel.Model,
  tblCarsColour.Colour
FROM
  (
    tblCars
    LEFT JOIN tblCarsModel ON tblCars.ID = tblCarsModel.ID
  )
  RIGHT JOIN tblCarsColour ON tblCarsColour.ID = tblCars.ID;
