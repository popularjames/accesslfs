SELECT Link_Table_Config.Location, Link_Table_Config.Server, Link_Table_Config.Database, Link_Table_Config.Table
FROM Link_Table_Config LEFT JOIN MSysObjects ON Link_Table_Config.Table = MSysObjects.Name
WHERE (((MSysObjects.Id) Is Null));
