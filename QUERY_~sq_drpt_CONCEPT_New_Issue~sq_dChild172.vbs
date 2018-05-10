PARAMETERS __ConceptID Value, __PayerNameID Value;
SELECT DISTINCTROW *
FROM v_CONCEPT_ValidationSummary_Payer AS rpt_CONCEPT_New_Issue
WHERE (([__ConceptID] = Adj_ConceptID)) AND ([__PayerNameID] = PayerNameID);
