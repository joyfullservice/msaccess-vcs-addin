SELECT
  [tblCars (Archive)].Manufacturer,
  tblCarsModel.Model
FROM
  [tblCars (Archive)]
  INNER JOIN tblCarsModel ON [tblCars (Archive)].ID = tblCarsModel.ID;
