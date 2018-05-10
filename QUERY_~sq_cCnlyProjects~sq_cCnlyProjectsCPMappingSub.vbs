PARAMETERS __ProjectID Value;
SELECT DISTINCTROW *
FROM CnlyProjectsCPMapping AS CnlyProjects
WHERE ([__ProjectID] = ProjectID);
