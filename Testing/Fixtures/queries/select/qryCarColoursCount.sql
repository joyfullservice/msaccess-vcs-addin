SELECT DISTINCT
  qryCars.Manufacturer,
  qryCars.Year,
  qryCars.Model,
  Count(qryCars.Colour) AS Colours
FROM
  qryCars
WHERE
  (
    (
      (qryCars.Price) < 15000
    )
  )
GROUP BY
  qryCars.Manufacturer,
  qryCars.Year,
  qryCars.Model
HAVING
  (
    (
      (qryCars.Year) = 2011
    )
  )
ORDER BY
  qryCars.Manufacturer,
  qryCars.Model;
