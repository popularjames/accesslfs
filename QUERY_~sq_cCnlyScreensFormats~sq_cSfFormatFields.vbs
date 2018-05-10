PARAMETERS __ScreenID Value, __Src Value;
SELECT DISTINCTROW *
FROM CnlyScreensFieldFormats AS CnlyScreensFormats
WHERE (([__ScreenID] = ScreenID)) AND ([__Src] = RecordSource);
