SELECT WorkflowNextStateID, StateName AS [Next State], WorkflowStateName AS [Current State], WorkflowName, AuditID, NextStateID, WorkflowID
FROM vCPpAuditsWorkflowNextState
WHERE AllowOverride=1 AND WorkflowID= -2147483624 AND  AuditID in (-2,-3)
ORDER BY AuditID, WorkflowName;
