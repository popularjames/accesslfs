SELECT vCPpWeightedCreditUsers.UserName, vCPpWeightedCreditUsers.UserID
FROM vCPpWeightedCreditUsers
WHERE vCPpWeightedCreditUsers.AuditID=-1 And vCPpWeightedCreditUsers.UserTypeID=0
ORDER BY Username;
