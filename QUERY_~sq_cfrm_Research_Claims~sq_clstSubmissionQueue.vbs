SELECT Description, SS.SubmisionID, SS.CriteriaID, SS.UserID, SubmissionDate, ss.CompleteDate
FROM CONCEPT_CRITERIA_Submission AS SS INNER JOIN Criteria_Hdr AS HH ON ss.CriteriaID = hh.criteriaID;
