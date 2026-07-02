UPDATE
  (
    (
      tblCars
      INNER JOIN tblCarsModel ON tblCars.ID = tblCarsModel.ID
    )
    INNER JOIN tblCarsColour ON tblCarsModel.ID = tblCarsColour.ID
  )
  INNER JOIN tblCarsPrice ON tblCarsColour.ID = tblCarsPrice.ID
SET
  tblCarsPrice.Price = [Price] * 1.05
WHERE
  (
    (
      (tblCars.Manufacturer) = "Audi"
    )
    AND (
      (tblCars.Year) = 2011
    )
    AND (
      (tblCarsModel.Model) = "GS"
    )
  );
