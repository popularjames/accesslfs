PARAMETERS __ScreenID Value, __TxtListLevel1 Value;
SELECT DISTINCTROW *
FROM CnlyScreensListFields AS CnlyScreens
WHERE (([__ScreenID] = ScreenID)) AND ([__TxtListLevel1] = ListLevel);
