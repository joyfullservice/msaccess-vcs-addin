SELECT
  tblCars.ID,
  tblCars.Manufacturer,
  [Manufacturer] & "|" & [ID] AS [Key],
  1 AS [Action],
  tblCars.Year AS [Year],
  tblCars.Model AS [Names],
  tblCars.Colour AS [Number],
  "USD" AS [Currency]
FROM
  tblCars;
