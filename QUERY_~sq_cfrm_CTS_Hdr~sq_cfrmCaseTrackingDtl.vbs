PARAMETERS __CaseId Value;
SELECT DISTINCTROW *
FROM (SELECT CtsCaseHdr_HIST.CaseId, CtsCaseHdr_HIST.ProviderNum, CtsCaseHdr_HIST.ICN, CtsCaseHdr_HIST.SourceDesc, CtsCaseHdr_HIST.RootCauseDesc, CtsCaseHdr_HIST.DispositionDesc, CtsCaseHdr_HIST.CategoryDesc, CtsCaseHdr_HIST.SubCategoryDesc, CtsCaseHdr_HIST.ActionDesc, CtsCaseHdr_HIST.StatusDesc, CtsCaseHdr_HIST.NoteText, CtsCaseHdr_HIST.LastUpdate, CtsCaseHdr_HIST.LastUserId, CtsCaseHdr_HIST.SeqNo, CtsCaseHdr_HIST.AssignedTo, * FROM CtsCaseHdr_HIST ORDER BY CtsCaseHdr_HIST.LastUpdate DESC , CtsCaseHdr_HIST.SeqNo DESC)  AS frm_CTS_Hdr
WHERE ([__CaseId] = CaseId);
