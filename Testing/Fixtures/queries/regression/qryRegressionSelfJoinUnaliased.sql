SELECT
  tblCars.Manufacturer AS PrimaryMake,
  tblCars_1.Manufacturer AS SiblingMake,
  tblCars.Year
FROM
  tblCars
  INNER JOIN tblCars AS tblCars_1 ON tblCars.Year = tblCars_1.Year
WHERE
  tblCars.ID < tblCars_1.ID;
