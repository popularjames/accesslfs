SELECT ADMIN_User_Account.AccountID, ADMIN_Client_Account.AcctAbbrev, ADMIN_Client_Account.AcctDesc
FROM ADMIN_User_Account INNER JOIN ADMIN_Client_Account ON ADMIN_User_Account.AccountID=ADMIN_Client_Account.AccountID
WHERE (((ADMIN_User_Account.UserID)=getusername()));
