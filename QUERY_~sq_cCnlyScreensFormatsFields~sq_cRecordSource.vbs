SELECT MSysObjects.Name, Switch([Type]=4,"Table",[Type]=5,"Query") AS Tp
FROM MSysObjects
WHERE (((MSysObjects.Flags) Not In (3,64,8,10,2,-2147483648)) AND ((MSysObjects.Type) In (1,4,5)))
ORDER BY MSysObjects.Type, MSysObjects.Name;
