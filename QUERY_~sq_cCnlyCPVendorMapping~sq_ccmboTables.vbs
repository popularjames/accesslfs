SELECT MSysObjects.Name, MSysObjects.ForeignName, IIf(MSysObjects.Type=1 Or MSysObjects.Type=5 Or MSysObjects.Type=6,MSysObjects.Name,MSysObjects.Name & ' (Linked)') AS Tables, Switch([Type]=1,"Table",[Type]=4,"Table",[Type]=6,"Table",[Type]=5,"Query") AS T
FROM MSysObjects
WHERE (((MSysObjects.Name) Not Like '*~sq*' And (MSysObjects.Name) Not Like 'msys*') AND ((MSysObjects.Type)=1 Or (MSysObjects.Type)=4 Or (MSysObjects.Type)=6) AND ((MSysObjects.ParentId)=251658241))
ORDER BY Switch([Type]=1,"Table",[Type]=4,"Table",[Type]=6,"Table",[Type]=5,"Query") DESC , MSysObjects.Name;
