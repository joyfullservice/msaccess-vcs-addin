SELECT
  tblCars.Manufacturer
FROM
  tblCars
UNION ALL
SELECT
  tblVehicles.Manufacturer
FROM
  tblVehicles;
