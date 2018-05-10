SELECT *
FROM LETTER_Work_Queue
WHERE RowCreateDt >= #10/19/2010# and RowCreateDt < #10-20-2010#  AND Status in ('W','R')AND (LetterType = 'VADRA')
ORDER BY lettertype, cnlyProvID, letterreqdt;
