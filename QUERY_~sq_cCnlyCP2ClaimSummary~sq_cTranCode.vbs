SELECT CC.TranCode, CC.TranCodeText, (CC.TranCode & Space(5) & '(' & CC.TranCodeText & ')') AS TranText
FROM vCPpClaimsCodes AS CC LEFT JOIN (SELECT visibilityid, trancode, trantype FROM vCpuAuditsClaimCodeVisibility WHERE NZ(auditid,0)=-1 And trantype=0)  AS CV ON (CC.TranCode=CV.TranCode) AND (CC.TranType=CV.TranType)
WHERE CV.VisibilityID is NULL And CC.TranType=0
ORDER BY CC.TranCode;
