PARAMETERS __ConceptID Value, __PayerNameID Value;
SELECT DISTINCTROW *
FROM v_CONCEPT_NIRF_Universe_W_Edits AS rpt_CONCEPT_New_Issue_2014
WHERE (([__ConceptID] = ConceptID)) AND ([__PayerNameID] = PayerNameID);
