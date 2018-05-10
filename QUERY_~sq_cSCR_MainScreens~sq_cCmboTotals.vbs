SELECT CnlyScreensTotals.TotalID, IIf([Global]=True,"GLB: ","") & [TotalName] AS TheName
FROM CnlyScreensTotals
ORDER BY CnlyScreensTotals.Global DESC , CnlyScreensTotals.TotalName;
