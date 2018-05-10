PARAMETERS __AuditID Value;
SELECT DISTINCTROW *
FROM qScrClaimTotalsVendor AS CnlyScrClaimRptClaimSummary
WHERE ([__AuditID] = AuditID);
