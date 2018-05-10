UPDATE Link_Table_Config SET Link_Table_Config.[Database] = "HUMANA_" & Mid([Database],InStr([Database],"_")+1)
WHERE (((Link_Table_Config.Database) Like "*_AUDITORS_*"));
