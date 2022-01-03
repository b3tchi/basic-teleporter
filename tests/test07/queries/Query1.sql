SELECT t_LargeNr.*
, t_Table2.Field1
FROM t_Table2
 INNER JOIN t_LargeNr
 ON t_Table2.ID = t_LargeNr.ID;
