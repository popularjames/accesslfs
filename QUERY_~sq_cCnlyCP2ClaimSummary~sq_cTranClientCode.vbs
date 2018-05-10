SELECT CLTC.TranCode, CLTC.TranCodeText, (CLTC.TranCode & space(3) & '(' & CLTC.TranCodeText & ')') AS TranText
FROM vCpuClaimsLedgerTranTypes AS CLT INNER JOIN vCpuClaimsLedgerTranCodesClient AS CLTC ON CLT.TranType=CLTC.TranType
WHERE AuditID=-1 And CLTC.TranType=0
ORDER BY CLTC.TranCode;
