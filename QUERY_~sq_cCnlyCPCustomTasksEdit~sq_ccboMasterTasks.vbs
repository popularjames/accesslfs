SELECT TaskName AS [Audit Master Tasks], TaskID, Billable
FROM vCPuMasterTasks
WHERE InactiveDate IS NULL;
