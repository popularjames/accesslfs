SELECT ue.*
FROM (ADMIN_User_Profile AS up INNER JOIN ADMIN_User_Exception AS ue ON up.UserProfID=ue.UserProfID) INNER JOIN ADMIN_Profile AS ap ON up.ProfileID=ap.ProfileID;
