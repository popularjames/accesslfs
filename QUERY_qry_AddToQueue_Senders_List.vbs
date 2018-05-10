SELECT ADMIN_User_Group.GroupID AS ID, "Group: " & [GroupName] AS Name, "Group" AS Type, ADMIN_User_Group.AccountID
FROM ADMIN_User_Group
UNION SELECT ADMIN_User_Account.UserID AS ID, ADMIN_User_Account.UserID AS Name, "Auditor" AS Type, ADMIN_User_Account.AccountID
FROM ADMIN_User_Account
ORDER BY Type DESC , Name;
