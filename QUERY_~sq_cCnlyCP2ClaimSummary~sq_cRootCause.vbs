SELECT CC.TranCode, CC.TranCodeText, CC.RootName, (CC.TranCode & Space(5) & '(' & CC.TranCodeText & ')') AS TranText
FROM vCPpClaimsCodes AS CC LEFT JOIN (SELECT visibilityid, trancode, trantype FROM vCPuAuditsClaimCodeVisibility WHERE auditid=-1 and trantype=20)  AS CV ON (CC.TranCode=CV.TranCode) AND (CC.TranType=CV.TranType)
WHERE ((CC.TranType)=20) And (CV.VisibilityID Is Null)
ORDER BY CC.TranCode;
