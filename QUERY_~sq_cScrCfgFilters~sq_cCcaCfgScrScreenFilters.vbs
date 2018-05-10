PARAMETERS __ScreenID Value;
SELECT DISTINCTROW *
FROM CnlyScreensFilters AS ScrCfgFilters
WHERE ([__ScreenID] = ScreenID);
