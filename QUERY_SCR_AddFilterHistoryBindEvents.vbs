INSERT INTO SCR_ScreensEvents ( ScreenID, EventType, Function, Seq )
SELECT SCR_Screens.ScreenID, "Screen Refresh" AS EventType, "BindFilterHisory" AS Function, 0 AS Seq
FROM SCR_Screens
WHERE (((SCR_Screens.ScreenID) Not In (SELECT ScreenID from SCR_ScreensEvents Where Function = "BindFilterHisory")));
