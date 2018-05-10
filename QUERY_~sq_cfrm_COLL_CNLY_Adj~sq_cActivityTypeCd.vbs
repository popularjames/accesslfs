SELECT XREF_ActivityType.ActivityTypeCd, XREF_ActivityType.ActivityTypeShrtNm, XREF_ActivityType.ActivityTypeTxt
FROM XREF_ActivityType
WHERE (((XREF_ActivityType.ActivityTypeCd) Not In ('OTHER','INTEREST INV','UNKNOWN','ARSETUP','WITHHOLDING INV')) AND ((XREF_ActivityType.ActivityTypeDomainCd)='CNLY'))
ORDER BY XREF_ActivityType.ActivityTypeCd;
