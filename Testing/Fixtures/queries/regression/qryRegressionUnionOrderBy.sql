SELECT
  tblCars.Manufacturer
FROM
  tblCars
UNION
SELECT
  tblVehicles.Manufacturer
FROM
  tblVehicles
ORDER BY
  tblCars.Manufacturer;
