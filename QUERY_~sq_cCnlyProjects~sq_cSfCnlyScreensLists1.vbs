PARAMETERS __ProjectID Value;
SELECT DISTINCTROW *
FROM CnlyProjectsFields AS CnlyProjects
WHERE ([__ProjectID] = ProjectID);
