SELECT CDate(Format([ScannedDt],"mm-dd-yyyy")) AS SDate, SCANNING_Image_Log_Tmp.CnlyClaimNum, SCANNING_Image_Log_Tmp.ImageType, Count(SCANNING_Image_Log_Tmp.CnlyClaimNum) AS CountOfCnlyClaimNum
FROM SCANNING_Image_Log_Tmp
GROUP BY CDate(Format([ScannedDt],"mm-dd-yyyy")), SCANNING_Image_Log_Tmp.CnlyClaimNum, SCANNING_Image_Log_Tmp.ImageType
HAVING (((Count(SCANNING_Image_Log_Tmp.CnlyClaimNum))>1));
