SELECT CONCEPT_Hdr.ConceptID, CONCEPT_Hdr.MaxChartRequest, CONCEPT_Hdr.ConceptDesc, CONCEPT_Hdr.ConceptStatus, Left([reviewtype],1) AS ReviewType1, CONCEPT_Hdr.ConceptGroup, CONCEPT_Hdr.DataType, CONCEPT_Hdr.ConceptLevel
FROM CONCEPT_Hdr
WHERE (((CONCEPT_Hdr.ConceptStatus) In (250,360,380,400,401)) AND ((Left([reviewtype],1)) Like "C*"))
ORDER BY CONCEPT_Hdr.ConceptID;
