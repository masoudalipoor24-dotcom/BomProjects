SELECT BOMHeaderID, ComponentItemID, Count(*) AS Cnt
FROM tblBOMLines
GROUP BY BOMHeaderID, ComponentItemID
HAVING Count(*) > 1;
