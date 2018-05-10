PARAMETERS __ConceptID Value, __PayerNameID Value;
SELECT DISTINCTROW *
FROM (SELECT CONCEPT_Dtl_Codes.ConceptID, CONCEPT_Dtl_Codes.CodeTypeID, CONCEPT_Dtl_Codes.Code, CONCEPT_Dtl_Codes.Reference, CONCEPT_Dtl_Codes.Comments, CONCEPT_XREF_CodeType.CodeDesc, CONCEPT_Dtl_Codes.PayerNameId FROM (CONCEPT_Dtl_Codes INNER JOIN CONCEPT_Hdr ON CONCEPT_Dtl_Codes.ConceptID = CONCEPT_Hdr.ConceptID) INNER JOIN CONCEPT_XREF_CodeType ON CONCEPT_Dtl_Codes.CodeTypeID = CONCEPT_XREF_CodeType.CodeTypeID)  AS rpt_CONCEPT_New_Issue
WHERE (([__ConceptID] = ConceptID)) AND ([__PayerNameID] = PayerNameId);
