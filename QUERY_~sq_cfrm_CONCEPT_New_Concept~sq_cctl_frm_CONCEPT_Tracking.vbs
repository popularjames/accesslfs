PARAMETERS __ConceptID Value;
SELECT DISTINCTROW *
FROM (SELECT * FROM CONCEPT_TRACKING ORDER BY ConceptId, TrackDate, DateEntered)  AS frm_CONCEPT_New_Concept
WHERE ([__ConceptID] = ConceptID);
