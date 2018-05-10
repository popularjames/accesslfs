SELECT DRG, MSDRGDesc, SUm(1) AS [Count]
FROM v_QUEUE_AuditorAssign_Claims
WHERE FromAuditor IN ( 'Theresa.Warren')
GROUP BY DRG, MSDRGDesc
ORDER BY drg;
