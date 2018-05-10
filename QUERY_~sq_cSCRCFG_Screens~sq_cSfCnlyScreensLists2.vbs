PARAMETERS __ScreenID Value, __TxtListLevel2 Value;
SELECT DISTINCTROW *
FROM SCR_ScreensListFields AS SCRCFG_Screens
WHERE (([__ScreenID] = ScreenID)) AND ([__TxtListLevel2] = ListLevel);
