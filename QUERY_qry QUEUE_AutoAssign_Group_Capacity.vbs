SELECT gc.GroupName, gc.AgeGroup, gc.Productivity, gc.Incomplete, Count(qau.CnlyClaimNum) AS TotalAvailable, gc.CarryOverCapacity, gc.CummulCarryOverCapacity, gc.NewCapacity, gc.Assigned, gc.LeftOverCapacity, gc.AgeGroupFillFactor
FROM QUEUE_AutoAssign_GroupCapacity AS gc INNER JOIN QUEUE_AutoAssign_UnAssignedClaims AS qau ON gc.AgeGroup = qau.AgeGroup
GROUP BY gc.GroupName, gc.AgeGroup, gc.Productivity, gc.Incomplete, gc.CarryOverCapacity, gc.CummulCarryOverCapacity, gc.NewCapacity, gc.Assigned, gc.LeftOverCapacity, gc.AgeGroupFillFactor;
