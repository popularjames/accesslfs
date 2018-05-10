PARAMETERS __ScreenID Value;
SELECT DISTINCTROW *
FROM CnlyScreensEvents AS CnlyScreens
WHERE ([__ScreenID] = ScreenID);
