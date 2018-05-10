SELECT MSysObjects.Name, Switch([Type]=4,"Table",[Type]=5,"Query") AS Tp
FROM MSysObjects
WHERE (((MSysObjects.Type) In (4,5)) AND ((MSysObjects.Flags) Not In (3,64)))
ORDER BY MSysObjects.Type, MSysObjects.Name;
