SELECT CnlyScreens.ScreenID, CnlyScreens.ScreenName, CnlyScreens.Included, CnlyScreens.Sort
FROM CnlyScreens
ORDER BY CnlyScreens.Sort, CnlyScreens.ScreenName;
