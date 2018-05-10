SELECT qry_TASK01.AuditId, qry_TASK01.TaskId, qry_TASK01.TaskName, qry_TASK01.TaskDesc, qry_TASK01.Requestor2 AS Requestor, qry_TASK01.AssignedTo2 AS AssignedTo, qry_TASK01.TaskWeight, qry_TASK01.RequestDt, qry_TASK01.RequestedDeliveryDt, qry_TASK01.ProjectedDeliveryDt, qry_TASK01.ActualDeliveryDt, qry_TASK01.CommentDt, qry_TASK01.CommentBy2 AS CommentBy, qry_TASK01.Comment, IIf([ActualDeliveryDt]="","No","Yes") AS Complete, DateDiff("d",[requestdt],Now()) AS TaskAge
FROM qry_TASK01
ORDER BY qry_TASK01.TaskId;
