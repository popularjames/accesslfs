PARAMETERS __ScreenID Value;
SELECT DISTINCTROW *
FROM (SELECT CnlyScreensNotes.* FROM CnlyScreensNotes ORDER BY CnlyScreensNotes.NoteDate DESC)  AS CnlyScreens
WHERE ([__ScreenID] = ScreenID);
