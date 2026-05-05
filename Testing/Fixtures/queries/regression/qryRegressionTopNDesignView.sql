SELECT
  TOP 3 tblCars.Manufacturer,
  tblCars.Year,
  tblCars.ID
FROM
  tblCars
ORDER BY
  tblCars.Year DESC;
