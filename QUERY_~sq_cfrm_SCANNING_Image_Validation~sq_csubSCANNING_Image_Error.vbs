SELECT ScannedDt, CnlyClaimNum, ErrMsg, PageCnt, ImageType, '' AS ImagePath, ImageName, cnlyProvID, ReceivedDt, ReceivedMeth, ScanOperator
FROM SCANNING_Image_Log_Tmp
WHERE 1=2;
