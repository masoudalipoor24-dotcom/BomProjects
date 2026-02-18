PARAMETERS pFGItemID Long;
SELECT h.*
FROM tblBOMHeader AS h
WHERE h.FGItemID = [pFGItemID]
  AND h.IsActive = True;


SELECT BOMHeaderID, ComponentItemID, Count(*) AS Cnt
FROM tblBOMLines
GROUP BY BOMHeaderID, ComponentItemID
HAVING Count(*) > 1;


SELECT
    l.BOMLineID,
    l.BOMHeaderID,
    l.LineNo,
    l.ComponentItemID,
    i.ItemCode        AS ComponentCode,
    i.ItemDescription AS ComponentDescription,
    l.QtyPer,
    l.UOMID,
    u.UOMCode,
    l.ScrapPct,
    l.EffectiveFrom,
    l.EffectiveTo,
    l.Notes
FROM (tblBOMLines AS l
      INNER JOIN tblItems AS i
      ON l.ComponentItemID = i.ItemID)
LEFT JOIN tblUOM AS u
ON l.UOMID = u.UOMID;


SELECT
    e.RootBOMHeaderID,
    e.RootItemID,
    e.ParentItemID,
    e.ComponentItemID,
    e.LevelNo,
    e.LineNo,
    e.QtyPer,
    e.ScrapPct,
    e.QtyWithScrap,
    e.ExtQty,
    u.UOMCode,
    ci.ItemCode AS ComponentCode,
    ci.ItemDescription AS ComponentDescription,
    e.SortKey,
    e.PathCodes
FROM (tmpBOMExplosion AS e
      INNER JOIN tblItems AS ci
      ON e.ComponentItemID = ci.ItemID)
LEFT JOIN tblUOM AS u
ON e.UOMID = u.UOMID
WHERE e.RunID = TempVars!BOMRunID
ORDER BY e.SortKey;


SELECT
    e.RootItemID,
    Sum(e.ExtQty * Nz(i.StdCost,0)) AS TotalStdCost
FROM tmpBOMExplosion AS e
INNER JOIN tblItems AS i
ON e.ComponentItemID = i.ItemID
WHERE e.RunID = TempVars!BOMRunID
GROUP BY e.RootItemID;


SELECT
    h.BOMHeaderID,
    fg.ItemCode AS FGCode,
    fg.ItemDescription AS FGDescription,
    h.VersionLabel,
    h.IsActive,
    e.LevelNo,
    e.SortKey,
    e.LineNo,
    ci.ItemCode AS ComponentCode,
    ci.ItemDescription AS ComponentDescription,
    e.QtyPer,
    e.ScrapPct,
    e.ExtQty,
    u.UOMCode
FROM (((tblBOMHeader AS h
INNER JOIN tblItems AS fg ON h.FGItemID = fg.ItemID)
INNER JOIN tmpBOMExplosion AS e ON h.BOMHeaderID = e.RootBOMHeaderID)
INNER JOIN tblItems AS ci ON e.ComponentItemID = ci.ItemID)
LEFT JOIN tblUOM AS u ON e.UOMID = u.UOMID
WHERE e.RunID = TempVars!BOMRunID
ORDER BY e.SortKey;
