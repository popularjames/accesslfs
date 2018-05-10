PARAMETERS __ProjectID Value;
SELECT DISTINCTROW *
FROM (SELECT * FROM CnlyProjectsWorkBars ORDER BY CnlyProjectsWorkBars.Sort)  AS CnlyProjects
WHERE ([__ProjectID] = ProjectID);
