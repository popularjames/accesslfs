SELECT U.UserName
FROM vCPuUsers AS U INNER JOIN vCPuAudits AS A ON U.CompanyID=A.CompanyID
WHERE ((U.UserTypeID=2 OR U.UserTypeID=6) AND (A.AuditID=-1))
ORDER BY UserName;
