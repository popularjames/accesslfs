SELECT CC.TranCode, CC.TranCodeText, CC.RootName, (CC.TranCode & Space(5) & '(' & CC.TranCodeText & ')') AS TranText
FROM vCPpClaimsCodes AS CC LEFT JOIN (SELECT visibilityid, trancode, trantype FROM vCPuAuditsClaimCodeVisibility WHERE auditid=1756 and trantype=10)  AS CV ON (CC.TranType=CV.TranType) AND (CC.TranCode=CV.TranCode)
WHERE ((CC.TranType)=10) And (CV.VisibilityID Is Null)
ORDER BY CC.TranCode;
