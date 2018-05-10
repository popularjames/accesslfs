PARAMETERS __ConceptID Value;
SELECT DISTINCTROW *
FROM (SELECT 
U.ConceptId
, U.ConceptState
, StateName
, SUM(U.ClaimCount) AS ClaimCount
, SUM(U.ClaimValue) AS ClaimValue
, U.DataType

FROM 
v_CONCEPT_NIRF_Financial_Universe U
GROUP BY 
U.ConceptId
, U.ConceptState
, U.StateName
, U.DataType)  AS rpt_CONCEPT_New_Issue_CMS_ONLY_2014
WHERE ([__ConceptID] = ConceptID);
