PARAMETERS __ScreenID Value;
SELECT DISTINCTROW *
FROM (SELECT * FROM CnlyScreensTabs)  AS CnlyScreens
WHERE ([__ScreenID] = ScreenID);
