PARAMETERS [Start Date] DateTime, [End Date] DateTime;
SELECT SCANNING_Image_Log.ScannedDt, SCANNING_Image_Log.CnlyClaimNum, SCANNING_Image_Log.ICN, SCANNING_Image_Log.ReceivedDt, SCANNING_Image_Log.ProvNum, SCANNING_Image_Log.ImageName, SCANNING_Image_Log.ValidationDt, CDate(Format([scanneddt],"mm-dd-yyyy")) AS Scanned_Dt
FROM SCANNING_Image_Log
WHERE (((SCANNING_Image_Log.ScannedDt) Between CDate([Start Date]) And DateAdd("d",1,CDate([End Date]))) AND ((SCANNING_Image_Log.ValidationDt) Is Not Null And (SCANNING_Image_Log.ValidationDt)<>#1/1/1900#));
