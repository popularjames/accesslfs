SELECT TranCode, TranCodeText, RootName, (TranCode & Space(5) & '(' & TranCodeText & ')') AS TranText
FROM vCPbClaimsCodes
WHERE ((TranType)=10)
ORDER BY TranCode;
