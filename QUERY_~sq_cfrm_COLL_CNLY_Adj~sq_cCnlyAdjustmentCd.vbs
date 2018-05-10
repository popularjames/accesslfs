SELECT XREF_CnlyAdjustment.CnlyAdjustmentCd, XREF_CnlyAdjustment.CnlyAdjustmentShrtNm, XREF_CnlyAdjustment.CnlyAdjustmentTxt, XREF_CnlyAdjustment.RequiredApprovalLevelCd
FROM XREF_CnlyAdjustment
WHERE (((XREF_CnlyAdjustment.CnlyAdjustmentDomainCd)='COLLECTION'));
