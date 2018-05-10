PARAMETERS __Rpt Value;
SELECT DISTINCTROW *
FROM (SELECT CnlyDtDupCriteriaDiff.FieldName, CnlyDtDupCriteriaDiff.Rpt, CnlyDtDupCriteriaDiff.FieldSeq FROM CnlyDtDupCriteriaDiff ORDER BY CnlyDtDupCriteriaDiff.FieldSeq)  AS CnlyDtCustomCriteriaSelection
WHERE ([__Rpt] = Rpt);
