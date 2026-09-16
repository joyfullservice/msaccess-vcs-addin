SELECT
  tblCars.Manufacturer & ' - ' & tblCars.Year AS DisplayText,
  'left = right' AS EqualText,
  'it''s "A - B"' AS EscapedText
FROM
  tblCars;
