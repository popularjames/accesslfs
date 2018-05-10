SELECT CnlyScreens.ScreenID, CnlyScreens.ScreenName
FROM CnlyScreens
WHERE (((CnlyScreens.Included)=True))
ORDER BY CnlyScreens.ScreenName;
