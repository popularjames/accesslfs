SELECT LETTER_Type.LetterType, LETTER_Type.LetterDesc, Count(QUEUE_Hdr.CnlyClaimNum) AS ClaimCnt, LETTER_Type.AccountID, LETTER_Type.ContractId
FROM (LETTER_Type INNER JOIN AUDITCLM_Process_Logics ON (LETTER_Type.AccountID = AUDITCLM_Process_Logics.AccountID) AND (LETTER_Type.LetterType = AUDITCLM_Process_Logics.DataType)) INNER JOIN QUEUE_Hdr ON (AUDITCLM_Process_Logics.AccountID = QUEUE_Hdr.AccountID) AND (AUDITCLM_Process_Logics.CurrQueue = QUEUE_Hdr.QueueType)
GROUP BY LETTER_Type.LetterType, LETTER_Type.LetterDesc, LETTER_Type.AccountID, LETTER_Type.ContractId
ORDER BY LETTER_Type.LetterType;
