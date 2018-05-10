SELECT cnlyCPClaimMapping.CPFieldName, cnlyCPClaimMapping.FieldName, cnlyCPClaimMapping.Keys, cnlyCPClaimMapping.Function, cnlyCPClaimMapping.cnlyCPClaimMappingID
FROM cnlyCPClaimMapping
WHERE cnlyCPClaimMapping.ScreenID=0
ORDER BY CPFieldName;
