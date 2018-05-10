PARAMETERS __ProjectID Value;
SELECT DISTINCTROW *
FROM (SELECT CnlyProjectsNotes.* FROM CnlyProjectsNotes ORDER BY CnlyProjectsNotes.NoteDate DESC)  AS CnlyProjects
WHERE ([__ProjectID] = ProjectID);
