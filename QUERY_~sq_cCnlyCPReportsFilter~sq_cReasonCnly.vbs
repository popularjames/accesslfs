SELECT CLTC.TranCode, CLTC.TranCodeText, (CLTC.TranCode & Space(5) & '(' & CLTC.TranCodeText & ')') AS TranText
FROM (vCpuClaimsLedgerTranCodesTypes AS CLTT INNER JOIN vCpuClaimsLedgerTranCodes AS CLTC ON CLTT.TranCode=CLTC.TranCode) LEFT JOIN vCpuAuditsClaimCodeVisibility AS ACCV ON (CLTT.TranCode=ACCV.TranCode) AND (CLTT.TranType=ACCV.TranType)
WHERE (((CLTT.TranType)=0) And ((Nz(AuditID,0))<>-37))
ORDER BY CLTC.TranCode;
