PARAMETERS __ScreenID Value, __TxtListLevel3 Value;
SELECT DISTINCTROW *
FROM CnlyScreensListFields AS CnlyScreens
WHERE (([__ScreenID] = ScreenID)) AND ([__TxtListLevel3] = ListLevel);
