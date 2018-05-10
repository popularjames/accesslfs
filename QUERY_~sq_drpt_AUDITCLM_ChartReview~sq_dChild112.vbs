PARAMETERS __cnlyclaimnum Value;
SELECT DISTINCTROW *
FROM AUDITCLM_Proc AS rpt_AUDITCLM_ChartReview
WHERE ([__cnlyclaimnum] = cnlyclaimnum);
