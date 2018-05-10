SELECT *
FROM (SELECT ProjectID, 0 as TypeID, "Main Grid" as Type , PrimaryRecordSource & " (Main Grid)" As Display FROM CnlyProjects    UNION SELECT ProjectID, WorkBarID as TypeID,"List" as Type, ListName & " (WorkBar)" As Display FROM CnlyProjectsWorkBars   UNION SELECT ProjectID, TabID as TypeID,  "Tab" as Type , Feature & "(Tab)" As Display FROM CnlyProjectsTabs)  AS rs
WHERE (((rs.ProjectID)=1))
ORDER BY rs.TypeID DESC;
