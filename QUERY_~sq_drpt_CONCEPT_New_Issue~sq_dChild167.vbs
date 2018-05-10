PARAMETERS __ConceptID Value, __PayerNameID Value;
SELECT DISTINCTROW *
FROM v_CONCEPT_NIRF_Dtl_State_Value_Rpt AS rpt_CONCEPT_New_Issue
WHERE (([__ConceptID] = ConceptID)) AND ([__PayerNameID] = PayerNameID);
