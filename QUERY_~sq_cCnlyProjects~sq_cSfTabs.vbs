PARAMETERS __ProjectID Value;
SELECT DISTINCTROW *
FROM (SELECT * FROM CnlyProjectsTabs)  AS CnlyProjects
WHERE ([__ProjectID] = ProjectID);
