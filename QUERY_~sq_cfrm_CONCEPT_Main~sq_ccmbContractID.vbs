SELECT ContractId, ContractDesc FROM ADMIN_Client_Contract
UNION SELECT TOP 1 0, 'All Contracts' FROM ADMIN_Client_Contract WHERE 1 = 1
ORDER BY 2;
