Attribute VB_Name = "modBOM"
Option Compare Database
Option Explicit

Private Const MAX_BOM_DEPTH As Long = 50

Public Function CreateGUID() As String
    Dim g As String
    g = CreateObject("Scriptlet.TypeLib").GUID
    g = Replace$(g, "{", "")
    g = Replace$(g, "}", "")
    CreateGUID = g
End Function

Public Function GetNextBOMVersionNo(ByVal FGItemID As Long) As Long
    Dim v As Variant
    v = DMax("VersionNo", "tblBOMHeader", "FGItemID=" & FGItemID)
    GetNextBOMVersionNo = Nz(v, 0) + 1
End Function

Public Function GetActiveBOMHeaderID(ByVal FGItemID As Long) As Variant
    GetActiveBOMHeaderID = DLookup("BOMHeaderID", "tblBOMHeader", "FGItemID=" & FGItemID & " AND IsActive=True")
End Function

Public Function ValidateNoCycle(ByVal ParentItemID As Long, ByVal ComponentItemID As Long, Optional ByVal AsOfDate As Date = 0) As Boolean
    If AsOfDate = 0 Then AsOfDate = Date

    If ComponentItemID = ParentItemID Then
        ValidateNoCycle = False
        Exit Function
    End If

    Dim visited As Object
    Set visited = CreateObject("Scripting.Dictionary")

    ValidateNoCycle = Not HasDescendant(ComponentItemID, ParentItemID, AsOfDate, visited, 0)
End Function

Private Function HasDescendant(ByVal StartItemID As Long, ByVal TargetItemID As Long, ByVal AsOfDate As Date, _
                               ByVal visited As Object, ByVal depth As Long) As Boolean
    If depth > MAX_BOM_DEPTH Then
        HasDescendant = True
        Exit Function
    End If

    If visited.Exists(CStr(StartItemID)) Then
        HasDescendant = False
        Exit Function
    End If
    visited.Add CStr(StartItemID), True

    Dim activeBOM As Variant
    activeBOM = GetActiveBOMHeaderID(StartItemID)
    If IsNull(activeBOM) Then
        HasDescendant = False
        Exit Function
    End If

    Dim db As DAO.Database
    Set db = CurrentDb()

    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT ComponentItemID " & _
          "FROM tblBOMLines " & _
          "WHERE BOMHeaderID=" & CLng(activeBOM) & " " & _
          "AND (EffectiveFrom Is Null OR EffectiveFrom<=" & SqlDateLiteral(AsOfDate) & ") " & _
          "AND (EffectiveTo Is Null OR EffectiveTo>=" & SqlDateLiteral(AsOfDate) & ");"

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    Do While Not rs.EOF
        Dim childID As Long
        childID = rs!ComponentItemID

        If childID = TargetItemID Then
            HasDescendant = True
            rs.Close
            Exit Function
        End If

        If HasDescendant(childID, TargetItemID, AsOfDate, visited, depth + 1) Then
            HasDescendant = True
            rs.Close
            Exit Function
        End If

        rs.MoveNext
    Loop
    rs.Close

    HasDescendant = False
End Function

Public Sub ActivateBOM(ByVal BOMHeaderID As Long, Optional ByVal AsOfDate As Date = 0)
    If AsOfDate = 0 Then AsOfDate = Date

    Dim db As DAO.Database
    Set db = CurrentDb()

    Dim rsH As DAO.Recordset
    Set rsH = db.OpenRecordset("SELECT BOMHeaderID, FGItemID FROM tblBOMHeader WHERE BOMHeaderID=" & BOMHeaderID, dbOpenSnapshot)
    If rsH.EOF Then
        rsH.Close
        Err.Raise vbObjectError + 100, "ActivateBOM", "Invalid BOMHeaderID."
    End If

    Dim fgID As Long
    fgID = rsH!FGItemID
    rsH.Close

    Dim rsL As DAO.Recordset
    Set rsL = db.OpenRecordset("SELECT ComponentItemID FROM tblBOMLines WHERE BOMHeaderID=" & BOMHeaderID & ";", dbOpenSnapshot)
    Do While Not rsL.EOF
        If Not ValidateNoCycle(fgID, rsL!ComponentItemID, AsOfDate) Then
            rsL.Close
            Err.Raise vbObjectError + 101, "ActivateBOM", "Activation blocked. Cycle detected."
        End If
        rsL.MoveNext
    Loop
    rsL.Close

    db.BeginTrans
    On Error GoTo EH

    db.Execute "UPDATE tblBOMHeader " & _
               "SET IsActive=False, ActiveKey=Null " & _
               "WHERE FGItemID=" & fgID & " AND BOMHeaderID<>" & BOMHeaderID & ";", dbFailOnError

    db.Execute "UPDATE tblBOMHeader " & _
               "SET IsActive=True, ActiveKey=FGItemID " & _
               "WHERE BOMHeaderID=" & BOMHeaderID & ";", dbFailOnError

    db.CommitTrans
    Exit Sub

EH:
    db.Rollback
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function CopyBOM(ByVal SourceBOMHeaderID As Long) As Long
    Dim db As DAO.Database
    Set db = CurrentDb()

    Dim rsS As DAO.Recordset
    Set rsS = db.OpenRecordset("SELECT FGItemID FROM tblBOMHeader WHERE BOMHeaderID=" & SourceBOMHeaderID, dbOpenSnapshot)
    If rsS.EOF Then
        rsS.Close
        Err.Raise vbObjectError + 200, "CopyBOM", "Source BOM not found."
    End If

    Dim fgID As Long
    fgID = rsS!FGItemID
    rsS.Close

    Dim nextVer As Long
    nextVer = GetNextBOMVersionNo(fgID)

    db.BeginTrans
    On Error GoTo EH

    db.Execute "INSERT INTO tblBOMHeader (FGItemID, VersionNo, VersionLabel, IsActive, ActiveKey, Notes, CreatedOn, CreatedBy) " & _
               "VALUES (" & fgID & ", " & nextVer & ", " & SqlTextLiteral("v" & nextVer) & ", False, Null, " & _
               SqlTextLiteral("Copied from BOMHeaderID=" & SourceBOMHeaderID) & ", Now(), " & SqlTextLiteral(Environ$("USERNAME")) & ");", dbFailOnError

    Dim newID As Long
    newID = db.OpenRecordset("SELECT @@IDENTITY AS NewID;", dbOpenSnapshot)!NewID

    db.Execute "INSERT INTO tblBOMLines (BOMHeaderID, LineNo, ComponentItemID, QtyPer, UOMID, ScrapPct, EffectiveFrom, EffectiveTo, Notes) " & _
               "SELECT " & newID & " AS BOMHeaderID, LineNo, ComponentItemID, QtyPer, UOMID, ScrapPct, EffectiveFrom, EffectiveTo, Notes " & _
               "FROM tblBOMLines WHERE BOMHeaderID=" & SourceBOMHeaderID & ";", dbFailOnError

    db.CommitTrans

    CopyBOM = newID
    Exit Function

EH:
    db.Rollback
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function BuildBOMExplosion(ByVal RootBOMHeaderID As Long, Optional ByVal AsOfDate As Date = 0) As String
    If AsOfDate = 0 Then AsOfDate = Date

    Dim db As DAO.Database
    Set db = CurrentDb()

    Dim runID As String
    runID = CreateGUID()

    Dim rsH As DAO.Recordset
    Set rsH = db.OpenRecordset("SELECT FGItemID FROM tblBOMHeader WHERE BOMHeaderID=" & RootBOMHeaderID, dbOpenSnapshot)
    If rsH.EOF Then
        rsH.Close
        Err.Raise vbObjectError + 300, "BuildBOMExplosion", "Invalid root BOMHeaderID."
    End If

    Dim rootItemID As Long
    rootItemID = rsH!FGItemID
    rsH.Close

    db.Execute "DELETE FROM tmpBOMExplosion WHERE CreatedOn < DateAdd('d', -7, Now());", dbFailOnError

    ExplodeNode db, runID, RootBOMHeaderID, rootItemID, rootItemID, 0, 1#, "", AsOfDate

    BuildBOMExplosion = runID
End Function

Private Sub ExplodeNode(ByVal db As DAO.Database, ByVal runID As String, ByVal RootBOMHeaderID As Long, _
                        ByVal rootItemID As Long, ByVal parentItemID As Long, ByVal levelNo As Long, _
                        ByVal parentExtQty As Double, ByVal sortPrefix As String, ByVal AsOfDate As Date)
    If levelNo > MAX_BOM_DEPTH Then Exit Sub

    Dim useBOMHeaderID As Variant

    If parentItemID = rootItemID And levelNo = 0 Then
        useBOMHeaderID = RootBOMHeaderID
    Else
        useBOMHeaderID = GetActiveBOMHeaderID(parentItemID)
        If IsNull(useBOMHeaderID) Then Exit Sub
    End If

    Dim sql As String
    sql = "SELECT LineNo, ComponentItemID, QtyPer, Nz(ScrapPct,0) AS ScrapPct, UOMID " & _
          "FROM tblBOMLines " & _
          "WHERE BOMHeaderID=" & CLng(useBOMHeaderID) & " " & _
          "AND (EffectiveFrom Is Null OR EffectiveFrom<=" & SqlDateLiteral(AsOfDate) & ") " & _
          "AND (EffectiveTo Is Null OR EffectiveTo>=" & SqlDateLiteral(AsOfDate) & ") " & _
          "ORDER BY LineNo;"

    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    Do While Not rs.EOF
        Dim lineNo As Long
        lineNo = Nz(rs!LineNo, 0)

        Dim compID As Long
        compID = rs!ComponentItemID

        Dim qtyPer As Double
        qtyPer = Nz(rs!QtyPer, 0)

        Dim scrapPct As Double
        scrapPct = Nz(rs!ScrapPct, 0)

        Dim qtyWithScrap As Double
        qtyWithScrap = qtyPer * (1# + (scrapPct / 100#))

        Dim extQty As Double
        extQty = parentExtQty * qtyWithScrap

        Dim sortKey As String
        sortKey = sortPrefix & Format$(lineNo, "0000") & "."

        db.Execute "INSERT INTO tmpBOMExplosion " & _
                   "(RunID, RootBOMHeaderID, RootItemID, ParentItemID, ComponentItemID, LevelNo, LineNo, " & _
                   "QtyPer, ScrapPct, QtyWithScrap, ExtQty, UOMID, SortKey, AsOfDate, CreatedOn, CreatedBy) " & _
                   "VALUES (" & _
                   SqlTextLiteral(runID) & ", " & RootBOMHeaderID & ", " & rootItemID & ", " & parentItemID & ", " & compID & ", " & (levelNo + 1) & ", " & lineNo & ", " & _
                   SqlDouble(qtyPer) & ", " & SqlDouble(scrapPct) & ", " & SqlDouble(qtyWithScrap) & ", " & SqlDouble(extQty) & ", " & SqlLongOrNull(rs!UOMID) & ", " & _
                   SqlTextLiteral(sortKey) & ", " & SqlDateLiteral(AsOfDate) & ", Now(), " & SqlTextLiteral(Environ$("USERNAME")) & ");", dbFailOnError

        If Not IsNull(GetActiveBOMHeaderID(compID)) Then
            ExplodeNode db, runID, RootBOMHeaderID, rootItemID, compID, levelNo + 1, extQty, sortKey, AsOfDate
        End If

        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Function SqlTextLiteral(ByVal value As String) As String
    SqlTextLiteral = "'" & Replace$(Nz(value, ""), "'", "''") & "'"
End Function

Private Function SqlDateLiteral(ByVal value As Date) As String
    SqlDateLiteral = "#" & Format$(value, "yyyy-mm-dd") & "#"
End Function

Private Function SqlDouble(ByVal value As Double) As String
    SqlDouble = Replace$(Trim$(Str$(value)), ",", ".")
End Function

Private Function SqlLongOrNull(ByVal value As Variant) As String
    If IsNull(value) Then
        SqlLongOrNull = "Null"
    Else
        SqlLongOrNull = CStr(CLng(value))
    End If
End Function
