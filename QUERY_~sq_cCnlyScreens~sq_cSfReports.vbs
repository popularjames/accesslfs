PARAMETERS __ScreenID Value;
SELECT DISTINCTROW *
FROM (SELECT * FROM CnlyScreensReports)  AS CnlyScreens
WHERE ([__ScreenID] = ScreenID);
