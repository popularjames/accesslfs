select UserID, OrderValue from (SELECT TOP 1 'View All' AS UserID, 1 AS OrderValue
FROM qFrmAdmin_Letters_Auditors)
UNION (Select UserID, 2 AS OrderValue
FROM qFrmAdmin_Letters_Auditors)
ORDER BY OrderValue, UserID;
