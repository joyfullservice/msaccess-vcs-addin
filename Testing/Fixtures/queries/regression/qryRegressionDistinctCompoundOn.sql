SELECT DISTINCT
  tblCars.Manufacturer,
  tblCarsModel.Model
FROM
  tblCars
  INNER JOIN tblCarsModel ON (tblCars.ID = tblCarsModel.ID)
  AND (tblCarsModel.ID > 0);
