PARAMETERS __ScreenID Value;
SELECT DISTINCTROW *
FROM (SELECT ScreenID,PrimaryRecordSource as Src FROM SCR_Screens UNION Select ScreenID, RecordSource From SCR_ScreensTabs  UNION  SELECT ScreenId, PrimaryListBoxRecordSource as Scr FROM SCR_Screens UNION  SELECT ScreenId, SecondaryListBoxRecordSource as Scr FROM SCR_Screens  UNION SELECT ScreenId, TertiaryListBoxRecordSource as Scr FROM SCR_Screens)  AS SCRCFG_Screens
WHERE ([__ScreenID] = ScreenID);
