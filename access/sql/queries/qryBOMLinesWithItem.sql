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
