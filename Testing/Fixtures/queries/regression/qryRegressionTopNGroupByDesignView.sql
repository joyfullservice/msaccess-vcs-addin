SELECT TOP 1
  tblCars.Manufacturer
FROM
  tblCars
  INNER JOIN tblCarsModel ON tblCars.ID = tblCarsModel.ID
GROUP BY
  tblCars.Manufacturer
ORDER BY
  tblCars.Manufacturer;
