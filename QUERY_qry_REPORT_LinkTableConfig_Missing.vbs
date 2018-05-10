SELECT "CMSPROD" AS Location, REPORT_Hdr.OutputTable AS [Table], "DS-FLD-009" AS Server, "CMS_Auditors_Reports" AS [Database]
FROM REPORT_Hdr
WHERE (((REPORT_Hdr.OutputTable) Is Not Null And (REPORT_Hdr.OutputTable) Not In (select nz(table,"--") from Link_Table_Config where location = "CMSPROD")));
