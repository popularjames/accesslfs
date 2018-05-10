PARAMETERS __ProjectID Value;
SELECT DISTINCTROW *
FROM (SELECT  ProjectID,  RecordSource as Src FROM CnlyProjectsTables UNION Select ProjectID, RecordSource From CnlyProjectsTabs)  AS CnlyProjects
WHERE ([__ProjectID] = ProjectID);
