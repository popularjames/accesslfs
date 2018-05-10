PARAMETERS __ProjectID Value;
SELECT DISTINCTROW *
FROM CnlyProjectsLog AS CnlyProjects
WHERE ([__ProjectID] = ProjectID);
