SELECT SCANNING_Image_Log_Tmp.ScannedDt, SCANNING_Image_Log_Tmp.CnlyClaimNum, SCANNING_Image_Log_Tmp.ImageType, SCANNING_Image_Log_Tmp.ImageName, "\\sn-philly-05\imaginge$\humana\dailyscans\" & [CnlyProvID] & "\" & [ImageName] & ".tif" AS TempPath, "\\sn-philly-05\imaginge$\humana\medicalrecords\" & [CnlyProvID] & "\" & [ImageName] & ".tif" AS MRPath
FROM [qry SCANNING: Identify duplicate scans per day] INNER JOIN SCANNING_Image_Log_Tmp ON ([qry SCANNING: Identify duplicate scans per day].ImageType=SCANNING_Image_Log_Tmp.ImageType) AND ([qry SCANNING: Identify duplicate scans per day].CnlyClaimNum=SCANNING_Image_Log_Tmp.CnlyClaimNum);