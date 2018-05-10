PARAMETERS __ConceptID Value, __PayerNameID Value;
SELECT DISTINCTROW *
FROM v_CONCEPT_NewIssueProposal_NEW_Dtl_State_Value_MANUAL AS rpt_CONCEPT_New_Issue_Manual
WHERE (([__ConceptID] = ConceptID)) AND ([__PayerNameID] = PayerNameID);
