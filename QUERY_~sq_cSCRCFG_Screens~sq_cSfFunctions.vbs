PARAMETERS __ScreenID Value;
SELECT DISTINCTROW *
FROM (SELECT * FROM SCR_ScreensFunctions)  AS SCRCFG_Screens
WHERE ([__ScreenID] = ScreenID);
