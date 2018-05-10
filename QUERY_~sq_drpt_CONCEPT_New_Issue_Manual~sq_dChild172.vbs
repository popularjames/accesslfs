PARAMETERS __ConceptID Value, __PayerNameID Value;
SELECT DISTINCTROW *
FROM CONCEPT_NIRF_Financials_Samples AS rpt_CONCEPT_New_Issue_Manual
WHERE (([__ConceptID] = ConceptID)) AND ([__PayerNameID] = PayerNameID);
