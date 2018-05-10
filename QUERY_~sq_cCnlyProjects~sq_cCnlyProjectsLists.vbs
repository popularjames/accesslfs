PARAMETERS __ProjectID Value;
SELECT DISTINCTROW *
FROM (SELECT * FROM CnlyProjectsLists ORDER BY ListName)  AS CnlyProjects
WHERE ([__ProjectID] = ProjectID);
