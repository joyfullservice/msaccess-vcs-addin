SELECT
  tblCars.Manufacturer,
  tblCars.Year,
  tblCars.ID
FROM
  tblCars
WHERE
  (
    (
      (tblCars.Manufacturer) In (
        SELECT
          [Manufacturer]
        FROM
          [tblCars] AS Tmp
        GROUP BY
          [Manufacturer]
        HAVING
          Count(*)> 1
      )
    )
  )
ORDER BY
  tblCars.Manufacturer;
