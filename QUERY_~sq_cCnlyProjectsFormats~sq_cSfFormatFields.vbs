PARAMETERS __ProjectID Value, __Src Value;
SELECT DISTINCTROW *
FROM CnlyProjectsFieldFormats AS CnlyProjectsFormats
WHERE (([__ProjectID] = ProjectID)) AND ([__Src] = RecordSource);
