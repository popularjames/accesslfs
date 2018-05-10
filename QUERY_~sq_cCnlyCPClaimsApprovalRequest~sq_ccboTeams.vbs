SELECT TeamName AS [Team Name], TeamID, Format(SpecifyApproverAtSend,'Yes/No') AS [Specify Approver At Send], DefaultDueInDays, TeamTypeID, TeamTypeName AS Type
FROM vCPpAuditsValidationClaimTeams AS AV
WHERE AV.AllowSendApprovalRequest=1 And AV.ClaimID='467CBDAB-9E1E-4BE6-9193-2E00220DE349';
