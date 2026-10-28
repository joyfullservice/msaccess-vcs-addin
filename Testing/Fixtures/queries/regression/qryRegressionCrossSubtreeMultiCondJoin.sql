SELECT
  tblCars.Manufacturer,
  tblCarsModel.Model,
  tblCarsColour.Colour,
  tblCarsPrice.Price
FROM
  (
    tblCars
    INNER JOIN (
      tblCarsModel
      INNER JOIN tblCarsColour ON tblCarsModel.ID = tblCarsColour.ID
    ) ON tblCars.ID = tblCarsModel.ID
  )
  INNER JOIN tblCarsPrice ON (
    tblCarsColour.ID = tblCarsPrice.ID
  )
  AND (tblCars.ID = tblCarsPrice.ID);
