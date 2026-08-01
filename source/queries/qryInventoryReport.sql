SELECT p.ProductName, c.CategoryName, p.UnitPrice, p.QuantityInStock,
       p.UnitPrice * p.QuantityInStock AS InventoryValue,
       IIf(p.QuantityInStock < 10, "REORDER", "OK") AS StockAlert
FROM tblProducts AS p
INNER JOIN tblCategories AS c ON p.CategoryID = c.CategoryID
ORDER BY IIf(p.QuantityInStock < 10, 0, 1), c.CategoryName, p.ProductName;