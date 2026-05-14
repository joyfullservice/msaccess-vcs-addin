PARAMETERS [Pick Date] DateTime;
SELECT
  tblCars.ID,
  tblCars.Manufacturer
FROM
  tblCars
WHERE
  (
    (
      (tblCars.Year) > Year([Pick Date])
    )
  );
