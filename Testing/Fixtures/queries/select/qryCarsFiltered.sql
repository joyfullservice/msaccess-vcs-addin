SELECT DISTINCT
  tblCars.Manufacturer,
  tblCars.Year,
  tblCarsModel.Model,
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
  (
    (
      (tblCars.Manufacturer)= "Ford"
    )
    AND (
      (tblCars.Year)= 2007
    )
    AND (
      (tblCarsModel.Model)= "Altima"
    )
    AND (
      (tblCarsColour.Colour)= "Orange"
    )
  );
