PARAMETERS __AuditID Value;
SELECT DISTINCTROW *
FROM qScrClaimTotalsAuditor AS CnlyScrClaimRptClaimSummary
WHERE ([__AuditID] = AuditID);
