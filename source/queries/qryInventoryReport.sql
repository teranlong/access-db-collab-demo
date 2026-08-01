SELECT p.ProductName, c.CategoryName, p.UnitPrice, p.QuantityInStock,
       p.UnitPrice * p.QuantityInStock AS InventoryValue
FROM tblProducts AS p
INNER JOIN tblCategories AS c ON p.CategoryID = c.CategoryID
ORDER BY c.CategoryName, p.ProductName;
