SELECT
  tblCars.ID,
  tblCars.Manufacturer,
  tblCarsModel.Model,
  tblCarsColour.Colour,
  carCopy.Manufacturer AS DuplicateManufacturer
FROM
  (
    (
      tblCars
      INNER JOIN tblCarsModel ON tblCars.ID = tblCarsModel.ID
    )
    INNER JOIN tblCarsColour ON tblCars.ID = tblCarsColour.ID
  )
  LEFT JOIN tblCars AS carCopy ON tblCars.Manufacturer = carCopy.Manufacturer;
