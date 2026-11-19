INSERT INTO tblArchive (
  ID,
  Name
) IN 'C:\path\to\other.accdb'
SELECT
  tblSource.ID,
  tblSource.Name
FROM
  tblSource;
