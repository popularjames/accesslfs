PARAMETERS __AuditID Value;
SELECT DISTINCTROW *
FROM qScrClaimTotalsStatus AS CnlyScrClaimRptClaimSummary
WHERE ([__AuditID] = AuditID);
