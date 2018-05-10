SELECT REPORT_Hdr.ReportID AS TabName, "frm_REPORT_Generic" AS FormName, "frm_REPORT_Main" AS AccessForm, REPORT_Hdr.ReportID AS FormValue
FROM REPORT_Hdr
WHERE (((REPORT_Hdr.ReportID) Not In (select nz(formvalue,"--") from qry_report_generaltabs)));
