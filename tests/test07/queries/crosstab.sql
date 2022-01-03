TRANSFORM Count(t_LargeNr.ID) AS CountOfID
SELECT t_LargeNr.ID
FROM t_LargeNr
GROUP BY t_LargeNr.ID
PIVOT t_LargeNr.LargeNr;
