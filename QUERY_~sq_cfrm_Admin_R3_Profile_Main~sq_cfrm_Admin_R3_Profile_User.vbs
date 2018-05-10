SELECT v_R3_Permission_ADMIN_User_Company.UserID, v_R3_Permission_ADMIN_User_Company.ProfileID, v_R3_Permission_ADMIN_User_Company.CompanyName, *
FROM v_R3_Permission_ADMIN_User_Company
WHERE (((v_R3_Permission_ADMIN_User_Company.UserID)=[cbUserID]));
