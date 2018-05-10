SELECT GroceryAdImages.City
FROM GroceryAdImages
WHERE Company='A_and_P'
GROUP BY GroceryAdImages.City
ORDER BY GroceryAdImages.City;
