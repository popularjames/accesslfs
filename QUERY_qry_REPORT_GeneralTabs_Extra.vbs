SELECT GENERAL_Tabs.RowID, GENERAL_Tabs.TabName, GENERAL_Tabs.FormName, GENERAL_Tabs.AccessForm, GENERAL_Tabs.FormValue
FROM GENERAL_Tabs
WHERE (((GENERAL_Tabs.TabName) Not In (select reportid from report_hdr)) AND ((GENERAL_Tabs.AccessForm)="frm_REPORT_Main"))
ORDER BY GENERAL_Tabs.TabName;
