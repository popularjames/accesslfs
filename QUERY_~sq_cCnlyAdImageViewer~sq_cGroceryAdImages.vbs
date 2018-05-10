SELECT [AutoID], Company, [WeekEndingDt], [FileName], Region, City
FROM GroceryAdImages
WHERE 1=1 And [Company]='A_and_P'
ORDER BY WeekEndingDt, FileName;
