SELECT DISTINCT CurPayerName
FROM QA_Review_Worktable
WHERE submitDate is null and AuditTeam = 'CNLY MN Team'
ORDER BY CurPayerName;
