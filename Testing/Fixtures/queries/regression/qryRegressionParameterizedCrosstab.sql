PARAMETERS [Select Manufacturer] Text (255);
TRANSFORM
  Count(qryCars.Colour) AS CountOfColour
SELECT
  qryCars.Manufacturer
FROM
  qryCars
WHERE
  (
    (
      (qryCars.Manufacturer) = [Select Manufacturer]
    )
  )
GROUP BY
  qryCars.Manufacturer
ORDER BY
  qryCars.Manufacturer
PIVOT
  qryCars.Model;
