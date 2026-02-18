SELECT
    e.RootItemID,
    Sum(e.ExtQty * Nz(i.StdCost,0)) AS TotalStdCost
FROM tmpBOMExplosion AS e
INNER JOIN tblItems AS i
ON e.ComponentItemID = i.ItemID
WHERE e.RunID = TempVars!BOMRunID
GROUP BY e.RootItemID;
