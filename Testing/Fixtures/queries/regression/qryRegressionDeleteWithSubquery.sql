DELETE tblCars.*
FROM
  tblCars
WHERE
  Exists (
    SELECT
      1
    FROM
      tblVehicles
    WHERE
      tblVehicles.Manufacturer = tblCars.Manufacturer
  ) = False;
