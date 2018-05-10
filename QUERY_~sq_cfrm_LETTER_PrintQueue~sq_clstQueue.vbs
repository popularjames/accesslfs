SELECT *
FROM LETTER_Work_Queue
WHERE RowCreateDt >= #8/24/2011# and RowCreateDt < #08-25-2011#
ORDER BY lettertype, cnlyProvID, letterreqdt;
