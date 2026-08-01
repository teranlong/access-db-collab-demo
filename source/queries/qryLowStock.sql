SELECT ProductName, QuantityInStock
FROM tblProducts
WHERE QuantityInStock < 10
ORDER BY QuantityInStock;
