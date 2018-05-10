PARAMETERS __AuditID Value;
SELECT DISTINCTROW *
FROM qScrClaimTotalsSource AS CnlyScrClaimRptClaimSummary
WHERE ([__AuditID] = AuditID);
