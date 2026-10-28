INSERT INTO tblCars (Manufacturer, [Year])
SELECT DISTINCT TOP 5 PERCENT
  tblVehicles.Manufacturer,
  2019 AS [Year]
FROM
  tblVehicles
  LEFT JOIN tblCars ON tblVehicles.Manufacturer = tblCars.Manufacturer
WHERE
  (
    (
      (tblCars.Manufacturer) Is Null
    )
  );
