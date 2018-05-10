SELECT [TeamTypeID], [TeamTypeName]
FROM vCPpAuditsValidationTeamType
WHERE AuditID=-1 And AllowApprovals=1
ORDER BY Ordinal;
