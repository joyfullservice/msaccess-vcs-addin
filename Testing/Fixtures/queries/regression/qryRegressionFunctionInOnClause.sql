SELECT
  cur.OrderID,
  cur.OrderDate
FROM
  tblOrders AS cur
  LEFT JOIN tblOrders AS prior ON (
    cur.CustomerID = prior.CustomerID
  )
  AND (
    prior.OrderDate = DateAdd('yyyy', -1, cur.OrderDate)
  );
