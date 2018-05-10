PARAMETERS __ScreenID Value;
SELECT DISTINCTROW *
FROM (SELECT * FROM CnlyScreensFunctions)  AS CnlyScreens
WHERE ([__ScreenID] = ScreenID);
