PARAMETERS __ClaimNum Value, __AuditID Value;
SELECT DISTINCTROW *
FROM vCpRptClaimNotes AS CnlyCpRptDetailNotesByAuditor
WHERE (([__ClaimNum] = ClaimNum)) AND ([__AuditID] = AuditID);
