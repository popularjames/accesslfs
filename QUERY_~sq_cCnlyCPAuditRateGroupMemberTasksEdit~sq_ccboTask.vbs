SELECT TaskID, TaskName, Billable
FROM vCPuMasterTasks
WHERE InactiveDate IS NULL
ORDER BY TaskName;
