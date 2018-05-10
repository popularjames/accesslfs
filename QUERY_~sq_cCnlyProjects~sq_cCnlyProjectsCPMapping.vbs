PARAMETERS __ProjectID Value;
SELECT DISTINCTROW *
FROM CnlyProjectsCpFields AS CnlyProjects
WHERE ([__ProjectID] = ProjectID);
