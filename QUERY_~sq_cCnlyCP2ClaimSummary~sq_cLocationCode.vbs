SELECT CC.TranCode, CC.TranCodeText, CC.RootName, (CC.TranCode & Space(5) & '(' & CC.TranCodeText & ')') AS TranText
FROM vCPpClaimsCodes AS CC LEFT JOIN (SELECT visibilityid, trancode, trantype FROM vCPuAuditsClaimCodeVisibility WHERE auditid=-1 and trantype=30)  AS CV ON (CC.TranType=CV.trantype) AND (CC.TranCode=CV.trancode)
WHERE (((CC.TranType)=20) AND ((CV.visibilityid) Is Null))
ORDER BY CC.TranCode;
