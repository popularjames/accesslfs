PARAMETERS __ScreenID Value;
SELECT DISTINCTROW *
FROM (SELECT * FROM CnlyScreensPowerBars)  AS CnlyScreens
WHERE ([__ScreenID] = ScreenID);
