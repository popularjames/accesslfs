SELECT MRReceived_Age, SUm(1) AS [Count]
FROM v_QUEUE_AuditorAssign_Claims
WHERE FromAuditor IN ( 'Theresa.Warren')
GROUP BY MRReceived_Age
ORDER BY MRReceived_Age;
