SELECT ADMIN_User_Profile_Audit_Hist.HistCreateDt, ADMIN_User_Profile_Audit_Hist.HistUser, ADMIN_User_Profile_Audit_Hist.UserProfID, ADMIN_User_Profile_Audit_Hist.UserID, ADMIN_User_Profile_Audit_Hist.ProfileID, ADMIN_User_Profile_Audit_Hist.HistImage
FROM ADMIN_User_Profile_Audit_Hist
WHERE (((ADMIN_User_Profile_Audit_Hist.UserID)=[cbUserID]) AND ((ADMIN_User_Profile_Audit_Hist.HistImage)="AFTER"))
ORDER BY ADMIN_User_Profile_Audit_Hist.HistCreateDt DESC , ADMIN_User_Profile_Audit_Hist.HistImage;
