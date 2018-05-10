SELECT vCPpAuditUsers.UserName, vCPpAuditUsers.UserID
FROM vCPpAuditUsers
WHERE vCPpAuditUsers.AuditID=-1 And vCPpAuditUsers.UserTypeID=0
ORDER BY Username;
