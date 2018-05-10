PARAMETERS __ScreenID Value, __TxtListLevel1 Value;
SELECT DISTINCTROW *
FROM SCR_ScreensListFields AS SCRCFG_Screens
WHERE (([__ScreenID] = ScreenID)) AND ([__TxtListLevel1] = ListLevel);
