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
