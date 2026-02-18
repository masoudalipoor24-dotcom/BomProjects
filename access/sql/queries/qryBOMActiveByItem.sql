PARAMETERS pFGItemID Long;
SELECT h.*
FROM tblBOMHeader AS h
WHERE h.FGItemID = [pFGItemID]
  AND h.IsActive = True;
