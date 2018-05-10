PARAMETERS __cnlyClaimNum Value;
SELECT DISTINCTROW *
FROM v_AUDITCLM_Diag AS rpt_AUDITCLM_ChartReview
WHERE ([__cnlyClaimNum] = cnlyClaimNum);
