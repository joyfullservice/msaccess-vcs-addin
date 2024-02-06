SELECT
  MSysQueries.Attribute,
  MSysQueries.Flag
FROM
  MSysQueries
WHERE
  (
    (
      (MSysQueries.Flag) In (
        SELECT
          Flag
        from
          [MSysQueries]
        where
          flag = 0
      )
    )
  );
