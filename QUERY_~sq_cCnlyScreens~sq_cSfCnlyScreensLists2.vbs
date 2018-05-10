PARAMETERS __ScreenID Value, __TxtListLevel2 Value;
SELECT DISTINCTROW *
FROM CnlyScreensListFields AS CnlyScreens
WHERE (([__ScreenID] = ScreenID)) AND ([__TxtListLevel2] = ListLevel);
