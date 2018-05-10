SELECT WorkflowID, WorkflowName, AuditID
FROM vCPuAuditsWorkflow
WHERE Active=0 AND  AuditID in (-2,-3)
ORDER BY AuditID, WorkflowName;
