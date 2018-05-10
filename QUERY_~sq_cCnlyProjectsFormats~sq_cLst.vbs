SELECT ProjectID,RecordSource as Src FROM CnlyProjectsTables Where ProjectID = 11 UNION Select ProjectID, RecordSource From CnlyProjectsTabs Where ProjectID = 11
ORDER BY Src;
