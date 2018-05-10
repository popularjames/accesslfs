PARAMETERS __AuditID Value;
SELECT DISTINCTROW *
FROM qScrClaimTotalsReason AS CnlyScrClaimRptClaimSummary
WHERE ([__AuditID] = AuditID);
