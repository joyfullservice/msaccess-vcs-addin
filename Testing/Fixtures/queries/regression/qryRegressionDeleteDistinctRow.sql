DELETE DISTINCTROW tblCars.*
FROM
  tblCars
WHERE
  (
    (
      (tblCars.ID)> 0
    )
  );
