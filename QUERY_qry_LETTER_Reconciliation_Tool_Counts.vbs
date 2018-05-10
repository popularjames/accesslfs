SELECT Count(1) AS CountofProvNum, qry_Letter_Reconciliation_Details.ProvNum
FROM qry_Letter_Reconciliation_Details
WHERE (((qry_Letter_Reconciliation_Details.CountOfRowNum)>1))
GROUP BY qry_Letter_Reconciliation_Details.ProvNum;
