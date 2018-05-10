PARAMETERS __ConceptID Value;
SELECT DISTINCTROW *
FROM v_CONCEPT_NIRF_DTL_Codes_Rpt AS rpt_CONCEPT_New_Issue_Alt
WHERE ([__ConceptID] = ConceptID);
