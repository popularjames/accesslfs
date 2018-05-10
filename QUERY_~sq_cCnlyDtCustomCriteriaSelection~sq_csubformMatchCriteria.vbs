PARAMETERS __Rpt Value;
SELECT DISTINCTROW *
FROM (SELECT CnlyDtDupCriteriaMatch.FieldName, CnlyDtDupCriteriaMatch.Rpt, CnlyDtDupCriteriaMatch.FieldSeq FROM CnlyDtDupCriteriaMatch ORDER BY CnlyDtDupCriteriaMatch.FieldSeq)  AS CnlyDtCustomCriteriaSelection
WHERE ([__Rpt] = Rpt);
