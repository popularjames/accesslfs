PARAMETERS __RowID Value;
SELECT DISTINCTROW *
FROM (SELECT GENERAL_Tabs_Linked_ProfileIDs.RowID, GENERAL_Tabs_Linked_ProfileIDs.ProfileID FROM GENERAL_Tabs_Linked_ProfileIDs)  AS frm_ADMIN_Report_sub1
WHERE ([__RowID] = RowID);
