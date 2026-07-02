SELECT
  tblCarsPrice.ID,
  Sum(tblCarsPrice.Price) AS TotalPrice
FROM
  (
    (
      tblCars
      INNER JOIN tblCarsModel ON tblCars.ID = tblCarsModel.ID
    )
    INNER JOIN tblCarsPrice ON (tblCars.Year = tblCarsPrice.Year)
    AND (
      tblCarsModel.ModelID = tblCarsPrice.ID
    )
  )
  INNER JOIN tblCarsColour ON tblCarsPrice.ID = tblCarsColour.ID
GROUP BY
  tblCarsPrice.ID;
