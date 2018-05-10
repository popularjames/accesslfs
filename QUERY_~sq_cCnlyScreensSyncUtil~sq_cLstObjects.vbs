SELECT CnlyScreensVersionsUtilities.UtiltiyID, CnlyScreensVersionsUtilities.UtilityName
FROM CnlyScreensVersionsUtilities
WHERE ((Left([UtilityName],4)="Sync"))
ORDER BY CnlyScreensVersionsUtilities.Sort, CnlyScreensVersionsUtilities.UtilityName;
