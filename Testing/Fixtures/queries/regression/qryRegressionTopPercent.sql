SELECT
  TOP 5 PERCENT tblCars.Manufacturer,
  tblCars.Year
FROM
  tblCars
ORDER BY
  tblCars.Year DESC;
