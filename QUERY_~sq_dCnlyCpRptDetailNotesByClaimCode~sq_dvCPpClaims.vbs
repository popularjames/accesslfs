PARAMETERS __ClaimNum Value, __AuditID Value;
SELECT DISTINCTROW *
FROM vCpRptClaimNotes AS CnlyCpRptDetailNotesByClaimCode
WHERE (([__ClaimNum] = ClaimNum)) AND ([__AuditID] = AuditID);
