SELECT DISTINCT S.State AS ConceptState, C.ConceptID, S.StateDesc AS StateName, G.EstAvailableClaims AS ClaimCount, G.EstAvailableDollars AS ClaimValue, NULL AS ClaimCountSample, Null AS ClaimValueState, C.DataType AS Reference, G.PayerNameID
FROM RPT_R0043C AS C INNER JOIN (RPT_R0043G AS G LEFT JOIN XREF_State AS S ON G.ProvStCd = S.State) ON C.ConceptID = G.ConceptID
WHERE G.EstAvailableClaims <> 0
AND G.EstAvailableDollars <> 0
AND C.ConceptID = 'CM_C1537';
