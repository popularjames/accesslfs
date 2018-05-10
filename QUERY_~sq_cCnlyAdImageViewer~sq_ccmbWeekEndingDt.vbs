SELECT GroceryAdImages.WeekEndingDt
FROM GroceryAdImages
WHERE Company='A_and_P'
GROUP BY GroceryAdImages.WeekEndingDt
ORDER BY GroceryAdImages.WeekEndingDt;
