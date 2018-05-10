PARAMETERS __ScreenID Value;
SELECT DISTINCTROW *
FROM (SELECT ScreenID,PrimaryRecordSource as Src FROM CnlyScreens UNION Select ScreenID, RecordSource From CnlyScreensTabs  UNION  SELECT ScreenId, PrimaryListBoxRecordSource as Scr FROM CnlyScreens UNION  SELECT ScreenId, SecondaryListBoxRecordSource as Scr FROM CnlyScreens  UNION SELECT ScreenId, TertiaryListBoxRecordSource as Scr FROM CnlyScreens)  AS CnlyScreens
WHERE ([__ScreenID] = ScreenID);
