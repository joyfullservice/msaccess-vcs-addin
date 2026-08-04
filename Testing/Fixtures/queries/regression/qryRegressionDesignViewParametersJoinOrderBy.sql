PARAMETERS [Min Qty] Long,
StatusFilter Short;
SELECT
  tblOrders.OrderID,
  tblCustomers.CustomerName,
  tblOrders.Quantity
FROM
  tblCustomers
  INNER JOIN tblOrders ON tblCustomers.ID = tblOrders.CustomerID
WHERE
  (
    (
      (
        tblOrders.Quantity >= [Min Qty]
        AND tblOrders.StatusID = StatusFilter
      )
    )
  )
ORDER BY
  tblOrders.Quantity,
  tblCustomers.CustomerName;
