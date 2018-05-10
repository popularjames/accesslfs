PARAMETERS __ScreenID Value;
SELECT DISTINCTROW *
FROM (SELECT SCR_ScreensNotes.* FROM SCR_ScreensNotes ORDER BY SCR_ScreensNotes.NoteDate DESC)  AS SCRCFG_Screens
WHERE ([__ScreenID] = ScreenID);
