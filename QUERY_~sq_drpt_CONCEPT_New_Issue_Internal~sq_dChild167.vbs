PARAMETERS __ConceptID Value;
SELECT DISTINCTROW *
FROM v_CONCEPT_NIRF_Dtl_State_Value_Rpt_Internal AS rpt_CONCEPT_New_Issue_Internal
WHERE ([__ConceptID] = ConceptID);
