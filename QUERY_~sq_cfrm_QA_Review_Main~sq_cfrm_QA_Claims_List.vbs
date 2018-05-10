SELECT TOP 1000 *
FROM v_QA_Review_WorkTable_Unsubmitted
WHERE AHclmstatus in ('314', '320', '320.2', '321', '322') and IcdversionCDflag = 9
ORDER BY Adj_ReviewType DESC , MRReceivedDt, AuditTeam, Auditor, ICN;
