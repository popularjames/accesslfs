SELECT AUF.FieldID, AUF.FieldName, AUF.FieldNum, AUF.Ordinal, AUF.DataType, DT.DataTypeName, AUF.Format, AUF.DefaultValue, AUF.Required, AUF.LimitToList
FROM vcpuAuditsUserFields AS AUF INNER JOIN vcpuDataTypes AS DT ON AUF.DataType=DT.DataType
WHERE (((AUF.AuditID)=-37))
ORDER BY AUF.FieldNum;
