PARAMETERS __ConceptID Value;
SELECT DISTINCTROW *
FROM v_CONCEPT_NIRF_Dtl_State_Value_Rpt AS rpt_CONCEPT_New_Issue_CMS_Only
WHERE ([__ConceptID] = ConceptID);
