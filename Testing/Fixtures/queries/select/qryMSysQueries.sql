SELECT
  MSysObjects.Name,
  MSysQueries.Attribute,
  MSysQueries.Flag,
  MSysQueries.Expression,
  MSysQueries.Name1,
  MSysQueries.Name2
FROM
  MSysObjects
  INNER JOIN MSysQueries ON MSysObjects.Id = MSysQueries.ObjectId
WHERE
  (
    (
      (MSysObjects.Flags)<> 3
    )
  )
ORDER BY
  MSysObjects.Name,
  MSysQueries.Attribute,
  MSysQueries.Flag;
