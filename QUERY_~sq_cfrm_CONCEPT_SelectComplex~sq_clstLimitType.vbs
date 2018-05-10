SELECT *
FROM (select distinct LimitID, LimitDesc from SELECT_Limits_Form where LimitID < 80 union select distinct 80, "Parameters" from select_limits_form union select distinct 0, "All" from select_limits_form)  AS a
ORDER BY limitdesc;
