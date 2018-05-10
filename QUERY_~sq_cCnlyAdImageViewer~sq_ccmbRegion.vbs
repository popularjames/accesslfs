SELECT GroceryAdImages.Region
FROM GroceryAdImages
WHERE Company='A_and_P'
GROUP BY GroceryAdImages.Region
ORDER BY GroceryAdImages.Region;
