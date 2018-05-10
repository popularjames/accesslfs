PARAMETERS __conceptid Value;
SELECT DISTINCTROW *
FROM (SELECT DISTINCT CONCEPT_Dtl_State.ConceptState, CONCEPT_Dtl_State.ConceptID, CONCEPT_XREF_State.StateName, CONCEPT_Dtl_State.PayerNameID FROM (CONCEPT_Dtl_State INNER JOIN CONCEPT_Hdr ON CONCEPT_Dtl_State.ConceptID = CONCEPT_Hdr.ConceptID) INNER JOIN CONCEPT_XREF_State ON CONCEPT_Dtl_State.ConceptState = CONCEPT_XREF_State.StateID)  AS rpt_CONCEPT_New_Issue_Alt
WHERE ([__conceptid] = conceptid);
