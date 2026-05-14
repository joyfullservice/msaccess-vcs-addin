UPDATE DISTINCTROW
  tblCars
SET
  tblCars.Colour = tblCars.Colour
WHERE
  (
    (
      (tblCars.ID)> 0
    )
  );
