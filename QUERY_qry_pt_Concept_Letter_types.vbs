SELECT LT.LetterType, LT.LetterDesc, LT.TemplateLoc, 
  (SELECT TOP 1
	RefLink FROM AUDITCLM_References R
	WHERE R.RefType = 'LETTER'
	AND 
	R.RefSubType = LT.LetterType
ORDER BY R.CreateDt DESC
  ) AS SampleDocPath
FROM Letter_Type AS LT
WHERE LT.AccountId = 1 AND LT.ADR = 1
ORDER BY LT.LetterType, LT.LetterDesc;