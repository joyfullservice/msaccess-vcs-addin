SELECT
  tblCars.ID,
  tblCars.Manufacturer INTO tblCarsExternal IN 'C:\Temp\VCSRegression.accdb'
FROM
  tblCars;
