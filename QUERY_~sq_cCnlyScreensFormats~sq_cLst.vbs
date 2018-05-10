SELECT ScreenID,PrimaryRecordSource as Src FROM CnlyScreens Where ScreenID = 44 UNION Select ScreenID, RecordSource From CnlyScreensTabs Where ScreenID = 44
ORDER BY Src;
