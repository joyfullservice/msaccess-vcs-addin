SELECT
  tblOrders.OrderID,
  tblProducts.ProductName,
  tblCategories.CategoryName,
  tblShippers.ShipperName
FROM
  (
    tblOrders
    INNER JOIN (
      (
        tblProducts
        INNER JOIN tblCategories ON tblProducts.CategoryID = tblCategories.CategoryID
      )
      INNER JOIN tblShippers ON tblProducts.ShipperID = tblShippers.ShipperID
    ) ON (
      tblOrders.ProductID = tblProducts.ProductID
    )
    AND (
      tblOrders.CategoryID = tblCategories.CategoryID
    )
  )
  INNER JOIN tblCustomers ON tblOrders.CustomerID = tblCustomers.CustomerID;
