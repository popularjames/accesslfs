SELECT Link_Table_Config.Location, Link_Table_Config.Table, Link_Table_Config.Server, Link_Table_Config.Database
FROM Link_Table_Config
WHERE location = "CMSPROD"
  AND (Table Like "*rpt*" Or Table Like "*report*")
  AND table not in  (select nz(outputtable,"--") from report_hdr)
  AND table not in ("REPORT_Hdr","REPORT_ListBox","REPORT_Parameter","REPORT_Runlog","v_AR_SETUP_Hdr_Error_Rpt","v_REPORT_LastRun");
