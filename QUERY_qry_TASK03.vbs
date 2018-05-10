SELECT qry_TASK01.TaskId, qry_TASK01.CommentDt, qry_TASK01.CommentBy2 AS CommentBy, qry_TASK01.Comment
FROM qry_TASK01
WHERE (((qry_TASK01.CommentDt)<>"") AND ((qry_TASK01.CommentBy2)<>"") AND ((qry_TASK01.Comment)<>""));
