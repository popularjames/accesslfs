SELECT DV.ICN, FWQ.*
FROM FAX_Work_Queue AS FWQ INNER JOIN Queue_RECON_Review_Results AS DV ON FWQ.CnlyClaimNum = DV.CnlyClaimNum
WHERE FWQ.Client_ext_Ref_ID IN ('1','4') AND DV.AssignedTo Like '*'
ORDER BY IIF(FWQ.ProcessedDate = #1/1/1900# ,#1/1/3000#,FWQ.ProcessedDate) DESC;
