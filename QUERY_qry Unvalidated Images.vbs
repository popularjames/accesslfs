SELECT SCANNING_Image_Log.ScannedDt, SCANNING_Image_Log.CnlyClaimNum, SCANNING_Image_Log.CnlyProvID, SCANNING_Image_Log.ImageType, SCANNING_Image_Log.PageCnt, SCANNING_Image_Log.PDFCnt, SCANNING_Image_Log.TIFCnt, SCANNING_Image_Log.ValidationDt
FROM SCANNING_Image_Log
WHERE (((SCANNING_Image_Log.ValidationDt) Is Null Or (SCANNING_Image_Log.ValidationDt)=#1/1/1900#));
