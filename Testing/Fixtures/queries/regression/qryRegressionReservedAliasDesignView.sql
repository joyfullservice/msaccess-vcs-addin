SELECT
  tblCars.ID,
  tblCars.Manufacturer AS [Option],
  tblCars.Model AS [First],
  tblCars.Colour AS [Second],
  tblCars.Year AS [To],
  [Manufacturer] & "|" & [ID] AS [Key],
  1 AS [Action],
  tblCars.Model AS [Names],
  tblCars.Colour AS [Number],
  "USD" AS [Currency],
  tblCars.Manufacturer AS Name
FROM
  tblCars;
