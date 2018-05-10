SELECT DISTINCT ReportingGroup AS Level1, hd.SubGroupSort, hd.ReportingSubGroup AS Level2, hd.ReportSort, hd.ReportName AS Level3, hd.ReportId, pr.ProfileID, gt.FormName, gt.AccessForm
FROM (Report_Hdr AS hd LEFT JOIN GENERAL_Tabs AS gt ON hd.ReportID = gt.FormValue) LEFT JOIN GENERAL_Tabs_Linked_ProfileIDs AS pr ON gt.RowID = pr.RowID
WHERE pr.ProfileID = "CM_Admin" AND gt.AccessForm = "frm_REPORT_Main" AND ActiveFlag = 1
ORDER BY 1, 2, 3, 4, 5;
