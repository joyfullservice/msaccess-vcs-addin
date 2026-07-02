TRANSFORM
  Count(qryCars.Colour) AS CountOfColour
SELECT
  qryCars.Manufacturer,
  qryCars.Year
FROM
  qryCars
WHERE
  (
    (
      (qryCars.Manufacturer) = "BMW"
      Or (qryCars.Manufacturer) = "Audi"
    )
  )
GROUP BY
  qryCars.Manufacturer,
  qryCars.Year
ORDER BY
  qryCars.Manufacturer,
  qryCars.Year
PIVOT
  qryCars.Model;
