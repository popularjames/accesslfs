SELECT ApprovalID, SentToEmail AS [Emailed To], TeamName AS [Team Name]
FROM vCPpClaimsApproval
WHERE ClaimID='{467CBDAB-9E1E-4BE6-9193-2E00220DE349}'
ORDER BY SentToEmail;
