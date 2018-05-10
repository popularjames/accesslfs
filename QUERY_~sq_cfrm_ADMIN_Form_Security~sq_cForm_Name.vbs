SELECT MSysObjects.Name
FROM MSysObjects
WHERE (((MSysObjects.Name) Not Like "~*") AND ((MSysObjects.Type)=-32768));
