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
