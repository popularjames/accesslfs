PARAMETERS __ConceptID Value;
SELECT DISTINCTROW *
FROM v_CONCEPT_ValidationSummary_Payer AS rpt_CONCEPT_New_Issue_CMS_Only
WHERE ([__ConceptID] = Adj_ConceptID);
