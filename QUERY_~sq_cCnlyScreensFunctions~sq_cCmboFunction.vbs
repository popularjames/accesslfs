SELECT O2.Name
FROM MSysObjects AS O1 INNER JOIN MSysObjects AS O2 ON O1.Id=O2.ParentId
WHERE (((O1.Name)="Forms"))
ORDER BY O2.Name;
