PARAMETERS __ProjectID Value;
SELECT DISTINCTROW *
FROM (SELECT * FROM CnlyProjectsReports)  AS CnlyProjects
WHERE ([__ProjectID] = ProjectID);
