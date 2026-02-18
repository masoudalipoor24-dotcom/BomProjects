Attribute VB_Name = "modBOMAutoSetup"
Option Compare Database
Option Explicit

Private mCurrentStep As String

Public Sub AutoSetupBOM(Optional ByVal RebuildQueries As Boolean = True)
    On Error GoTo EH

    Application.Echo False
    DoCmd.Hourglass True

    mCurrentStep = U("0627 06CC 062C 0627 062F 0020 062C 062F 0648 0644 200C 0647 0627")
    EnsureTables

    mCurrentStep = U("0627 06CC 062C 0627 062F 0020 0627 06CC 0646 062F 06A9 0633 200C 0647 0627")
    RunStepWithTolerance StepIndexes:=True

    mCurrentStep = U("0627 06CC 062C 0627 062F 0020 0631 0648 0627 0628 0637")
    RunStepWithTolerance StepRelationships:=True

    If RebuildQueries Then
        mCurrentStep = U("0627 06CC 062C 0627 062F 0020 06A9 0648 0626 0631 06CC 200C 0647 0627 06CC 0020 0630 062E 06CC 0631 0647 200C 0634 062F 0647")
        RunStepWithTolerance StepSavedQueries:=True
    End If

    mCurrentStep = U("0627 0639 0645 0627 0644 0020 0645 0642 062F 0627 0631 0020 067E 06CC 0634 200C 0641 0631 0636 0020 0648 0020 0642 0648 0627 0639 062F 0020 0627 0639 062A 0628 0627 0631 0633 0646 062C 06CC")
    RunStepWithTolerance StepApplyRules:=True

    mCurrentStep = U("062B 0628 062A 0020 062F 0627 062F 0647 200C 0647 0627 06CC 0020 067E 0627 06CC 0647")
    RunStepWithTolerance StepSeedData:=True

    mCurrentStep = U("067E 0627 06CC 0627 0646")
    DoCmd.Hourglass False
    Application.Echo True
    MsgBoxU U("0631 0627 0647 200C 0627 0646 062F 0627 0632 06CC 0020 062E 0648 062F 06A9 0627 0631 0020 0042 004F 004D 0020 0628 0627 0020 0645 0648 0641 0642 06CC 062A 0020 0627 0646 062C 0627 0645 0020 0634 062F 002E"), vbInformation
    Exit Sub

EH:
    Dim errNo As Long
    Dim errDesc As String
    errNo = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    DoCmd.Hourglass False
    Application.Echo True
    MsgBoxU U("0631 0627 0647 200C 0627 0646 062F 0627 0632 06CC 0020 062E 0648 062F 06A9 0627 0631 0020 0042 004F 004D 0020 062F 0631 0020 0645 0631 062D 0644 0647 0020 0632 06CC 0631 0020 0645 062A 0648 0642 0641 0020 0634 062F 003A 0020") & mCurrentStep & vbCrLf & _
           U("062E 0637 0627 0020") & errNo & ": " & errDesc, vbCritical
    Err.Clear
End Sub

Private Sub RunStepWithTolerance(Optional ByVal StepIndexes As Boolean = False, _
                                 Optional ByVal StepRelationships As Boolean = False, _
                                 Optional ByVal StepSavedQueries As Boolean = False, _
                                 Optional ByVal StepApplyRules As Boolean = False, _
                                 Optional ByVal StepSeedData As Boolean = False)
    On Error GoTo EH

    If StepIndexes Then
        EnsureIndexes
        Exit Sub
    End If

    If StepRelationships Then
        EnsureRelationships
        Exit Sub
    End If

    If StepSavedQueries Then
        EnsureSavedQueries
        Exit Sub
    End If

    If StepApplyRules Then
        ApplyBOMDefaultsAndRules
        Exit Sub
    End If

    If StepSeedData Then
        SeedLookupData
        Exit Sub
    End If
    Exit Sub

EH:
    If ShouldIgnoreSetupError(mCurrentStep, Err.Number, Err.Description) Then
        Err.Clear
        Exit Sub
    End If

    Err.Raise Err.Number, "RunStepWithTolerance(" & mCurrentStep & ")", Err.Description
End Sub

Public Sub AutoSetupBOMWithUI()
    On Error GoTo EH
    AutoSetupBOM True
    BuildBOMUIObjects
    DoCmd.OpenForm "frmMainMenu"
    MsgBoxU U("0631 0627 0647 200C 0627 0646 062F 0627 0632 06CC 0020 062E 0648 062F 06A9 0627 0631 0020 0042 004F 004D 0020 0648 0020 0631 0627 0628 0637 0020 06A9 0627 0631 0628 0631 06CC 0020 0628 0627 0020 0645 0648 0641 0642 06CC 062A 0020 0627 0646 062C 0627 0645 0020 0634 062F 002E"), vbInformation
    Exit Sub
EH:
    MsgBoxU U("0631 0627 0647 200C 0627 0646 062F 0627 0632 06CC 0020 062E 0648 062F 06A9 0627 0631 0020 0042 004F 004D 0020 0648 0020 0631 0627 0628 0637 0020 06A9 0627 0631 0628 0631 06CC 0020 0646 0627 0645 0648 0641 0642 0020 0628 0648 062F 003A 0020") & Err.Description, vbCritical
End Sub

Public Sub RepairPersianUI_NoLocaleChange()
    On Error GoTo EH

    AutoSetupBOM True
    ApplyBOMDefaultsAndRules
    SeedLookupData
    BuildBOMUIObjects

    DoCmd.OpenForm "frmMainMenu"
    MsgBoxU U("062A 0639 0645 06CC 0631 0020 0648 0020 0628 0627 0632 0633 0627 0632 06CC 0020 0646 0645 0627 06CC 0634 0020 0641 0627 0631 0633 06CC 0020 0628 062F 0648 0646 0020 062A 063A 06CC 06CC 0631 0020 062A 0646 0638 06CC 0645 0627 062A 0020 0633 06CC 0633 062A 0645 0020 0627 0646 062C 0627 0645 0020 0634 062F 002E"), vbInformation
    Exit Sub
EH:
    MsgBoxU U("062A 0639 0645 06CC 0631 0020 0646 0645 0627 06CC 0634 0020 0641 0627 0631 0633 06CC 0020 0646 0627 0645 0648 0641 0642 0020 0628 0648 062F 003A 0020") & Err.Description, vbCritical
End Sub

Public Function RunAutoSetupBOM() As Boolean
    On Error GoTo EH
    AutoSetupBOM True
    RunAutoSetupBOM = True
    Exit Function
EH:
    RunAutoSetupBOM = False
End Function

Private Sub EnsureTables()
    If Not TableExists("tblUOM") Then
        ExecDDL "CREATE TABLE tblUOM (" & _
                "UOMID AUTOINCREMENT CONSTRAINT PK_tblUOM PRIMARY KEY, " & _
                "UOMCode TEXT(10) NOT NULL, " & _
                "UOMName TEXT(50), " & _
                "IsActive YESNO NOT NULL" & _
                ");"
    End If

    If Not TableExists("tblItemType") Then
        ExecDDL "CREATE TABLE tblItemType (" & _
                "ItemTypeID AUTOINCREMENT CONSTRAINT PK_tblItemType PRIMARY KEY, " & _
                "TypeCode TEXT(10) NOT NULL, " & _
                "TypeName TEXT(50)" & _
                ");"
    End If

    If Not TableExists("tblItems") Then
        ExecDDL "CREATE TABLE tblItems (" & _
                "ItemID AUTOINCREMENT CONSTRAINT PK_tblItems PRIMARY KEY, " & _
                "ItemCode TEXT(30) NOT NULL, " & _
                "ItemDescription TEXT(255) NOT NULL, " & _
                "UOMID LONG, " & _
                "ItemTypeID LONG, " & _
                "IsActive YESNO NOT NULL, " & _
                "StdCost CURRENCY, " & _
                "CreatedOn DATETIME, " & _
                "CreatedBy TEXT(50), " & _
                "ModifiedOn DATETIME, " & _
                "ModifiedBy TEXT(50)" & _
                ");"
    End If

    If Not TableExists("tblBOMHeader") Then
        ExecDDL "CREATE TABLE tblBOMHeader (" & _
                "BOMHeaderID AUTOINCREMENT CONSTRAINT PK_tblBOMHeader PRIMARY KEY, " & _
                "FGItemID LONG NOT NULL, " & _
                "VersionNo LONG NOT NULL, " & _
                "VersionLabel TEXT(10), " & _
                "IsActive YESNO NOT NULL, " & _
                "ActiveKey LONG, " & _
                "Notes LONGTEXT, " & _
                "CreatedOn DATETIME, " & _
                "CreatedBy TEXT(50)" & _
                ");"
    End If

    If Not TableExists("tblBOMLines") Then
        ExecDDL "CREATE TABLE tblBOMLines (" & _
                "BOMLineID AUTOINCREMENT CONSTRAINT PK_tblBOMLines PRIMARY KEY, " & _
                "BOMHeaderID LONG NOT NULL, " & _
                "LineNo LONG, " & _
                "ComponentItemID LONG NOT NULL, " & _
                "QtyPer DOUBLE NOT NULL, " & _
                "UOMID LONG, " & _
                "ScrapPct DOUBLE, " & _
                "EffectiveFrom DATETIME, " & _
                "EffectiveTo DATETIME, " & _
                "Notes LONGTEXT" & _
                ");"
    End If

    If Not TableExists("tmpBOMExplosion") Then
        ExecDDL "CREATE TABLE tmpBOMExplosion (" & _
                "ExplosionID AUTOINCREMENT CONSTRAINT PK_tmpBOMExplosion PRIMARY KEY, " & _
                "RunID TEXT(36) NOT NULL, " & _
                "RootBOMHeaderID LONG NOT NULL, " & _
                "RootItemID LONG NOT NULL, " & _
                "ParentItemID LONG, " & _
                "ComponentItemID LONG NOT NULL, " & _
                "LevelNo LONG NOT NULL, " & _
                "LineNo LONG, " & _
                "QtyPer DOUBLE, " & _
                "ScrapPct DOUBLE, " & _
                "QtyWithScrap DOUBLE, " & _
                "ExtQty DOUBLE, " & _
                "UOMID LONG, " & _
                "SortKey TEXT(255), " & _
                "PathCodes LONGTEXT, " & _
                "AsOfDate DATETIME, " & _
                "CreatedOn DATETIME, " & _
                "CreatedBy TEXT(50)" & _
                ");"
    End If
End Sub

Private Sub EnsureIndexes()
    TryEnsureIndex "tblUOM", "UX_tblUOM_UOMCode", "UOMCode", True
    TryEnsureIndex "tblItemType", "UX_tblItemType_TypeCode", "TypeCode", True
    TryEnsureIndex "tblItems", "UX_tblItems_ItemCode", "ItemCode", True
    TryEnsureIndex "tblItems", "IX_tblItems_UOMID", "UOMID", False
    TryEnsureIndex "tblItems", "IX_tblItems_ItemTypeID", "ItemTypeID", False
    TryEnsureIndex "tblBOMHeader", "UX_tblBOMHeader_FG_Version", "FGItemID, VersionNo", True
    TryEnsureIndex "tblBOMHeader", "UX_tblBOMHeader_ActiveKey", "ActiveKey", True
    TryEnsureIndex "tblBOMHeader", "IX_tblBOMHeader_FGItemID", "FGItemID", False
    TryEnsureIndex "tblBOMLines", "UX_tblBOMLines_Header_Component", "BOMHeaderID, ComponentItemID", True
    TryEnsureIndex "tblBOMLines", "UX_tblBOMLines_Header_LineNo", "BOMHeaderID, LineNo", True
    TryEnsureIndex "tblBOMLines", "IX_tblBOMLines_ComponentItemID", "ComponentItemID", False
    TryEnsureIndex "tmpBOMExplosion", "IX_tmpBOMExplosion_RunID", "RunID", False
    TryEnsureIndex "tmpBOMExplosion", "IX_tmpBOMExplosion_Root", "RootBOMHeaderID", False
    mCurrentStep = U("0627 06CC 062C 0627 062F 0020 0627 06CC 0646 062F 06A9 0633 200C 0647 0627")
End Sub

Private Sub EnsureIndexWithStep(ByVal tableName As String, ByVal indexName As String, ByVal fieldsCsv As String, ByVal isUnique As Boolean)
    mCurrentStep = U("0627 06CC 062C 0627 062F 0020 0627 06CC 0646 062F 06A9 0633 003A 0020") & tableName & "." & indexName
    EnsureIndex tableName, indexName, fieldsCsv, isUnique
End Sub

Private Sub TryEnsureIndex(ByVal tableName As String, ByVal indexName As String, ByVal fieldsCsv As String, ByVal isUnique As Boolean)
    On Error GoTo EH
    EnsureIndexWithStep tableName, indexName, fieldsCsv, isUnique
    Exit Sub
EH:
    If IndexExists(tableName, indexName) Then
        Err.Clear
        Exit Sub
    End If
    If IsIndexErrorIgnorable(Err.Number, Err.Description) Then
        Err.Clear
        Exit Sub
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub EnsureRelationships()
    TryEnsureRelation "FK_tblItems_tblUOM", "tblUOM", "tblItems", "UOMID", "UOMID", False
    TryEnsureRelation "FK_tblItems_tblItemType", "tblItemType", "tblItems", "ItemTypeID", "ItemTypeID", False
    TryEnsureRelation "FK_tblBOMHeader_tblItems_FG", "tblItems", "tblBOMHeader", "ItemID", "FGItemID", False
    TryEnsureRelation "FK_tblBOMLines_tblBOMHeader", "tblBOMHeader", "tblBOMLines", "BOMHeaderID", "BOMHeaderID", True
    TryEnsureRelation "FK_tblBOMLines_tblItems_Component", "tblItems", "tblBOMLines", "ItemID", "ComponentItemID", False
    TryEnsureRelation "FK_tblBOMLines_tblUOM", "tblUOM", "tblBOMLines", "UOMID", "UOMID", False
End Sub

Private Sub TryEnsureRelation(ByVal relationName As String, ByVal parentTable As String, ByVal childTable As String, _
                              ByVal parentField As String, ByVal childField As String, _
                              Optional ByVal cascadeDelete As Boolean = False)
    On Error GoTo EH
    mCurrentStep = U("0627 06CC 062C 0627 062F 0020 0631 0627 0628 0637 0647 003A 0020") & relationName
    EnsureRelation relationName, parentTable, childTable, parentField, childField, cascadeDelete
    Exit Sub
EH:
    If IsAlreadyExistsError(Err.Number, Err.Description) Then
        Err.Clear
        Exit Sub
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub EnsureSavedQueries()
    SafeUpsertQueryDef "qryBOMActiveByItem", _
        "PARAMETERS pFGItemID Long;" & vbCrLf & _
        "SELECT h.*" & vbCrLf & _
        "FROM tblBOMHeader AS h" & vbCrLf & _
        "WHERE h.FGItemID = [pFGItemID]" & vbCrLf & _
        "  AND h.IsActive = True;"

    SafeUpsertQueryDef "qryValidationDuplicates", _
        "SELECT BOMHeaderID, ComponentItemID, Count(*) AS Cnt" & vbCrLf & _
        "FROM tblBOMLines" & vbCrLf & _
        "GROUP BY BOMHeaderID, ComponentItemID" & vbCrLf & _
        "HAVING Count(*) > 1;"

    SafeUpsertQueryDef "qryBOMLinesWithItem", _
        "SELECT" & vbCrLf & _
        "    l.BOMLineID," & vbCrLf & _
        "    l.BOMHeaderID," & vbCrLf & _
        "    l.LineNo," & vbCrLf & _
        "    l.ComponentItemID," & vbCrLf & _
        "    i.ItemCode        AS ComponentCode," & vbCrLf & _
        "    i.ItemDescription AS ComponentDescription," & vbCrLf & _
        "    l.QtyPer," & vbCrLf & _
        "    l.UOMID," & vbCrLf & _
        "    u.UOMCode," & vbCrLf & _
        "    l.ScrapPct," & vbCrLf & _
        "    l.EffectiveFrom," & vbCrLf & _
        "    l.EffectiveTo," & vbCrLf & _
        "    l.Notes" & vbCrLf & _
        "FROM (tblBOMLines AS l" & vbCrLf & _
        "      INNER JOIN tblItems AS i" & vbCrLf & _
        "      ON l.ComponentItemID = i.ItemID)" & vbCrLf & _
        "LEFT JOIN tblUOM AS u" & vbCrLf & _
        "ON l.UOMID = u.UOMID;"

    SafeUpsertQueryDef "qryBOMExplosion", _
        "SELECT" & vbCrLf & _
        "    e.RootBOMHeaderID," & vbCrLf & _
        "    e.RootItemID," & vbCrLf & _
        "    e.ParentItemID," & vbCrLf & _
        "    e.ComponentItemID," & vbCrLf & _
        "    e.LevelNo," & vbCrLf & _
        "    e.LineNo," & vbCrLf & _
        "    e.QtyPer," & vbCrLf & _
        "    e.ScrapPct," & vbCrLf & _
        "    e.QtyWithScrap," & vbCrLf & _
        "    e.ExtQty," & vbCrLf & _
        "    u.UOMCode," & vbCrLf & _
        "    ci.ItemCode AS ComponentCode," & vbCrLf & _
        "    ci.ItemDescription AS ComponentDescription," & vbCrLf & _
        "    e.SortKey," & vbCrLf & _
        "    e.PathCodes" & vbCrLf & _
        "FROM (tmpBOMExplosion AS e" & vbCrLf & _
        "      INNER JOIN tblItems AS ci" & vbCrLf & _
        "      ON e.ComponentItemID = ci.ItemID)" & vbCrLf & _
        "LEFT JOIN tblUOM AS u" & vbCrLf & _
        "ON e.UOMID = u.UOMID" & vbCrLf & _
        "WHERE e.RunID = TempVars!BOMRunID" & vbCrLf & _
        "ORDER BY e.SortKey;"

    SafeUpsertQueryDef "qryBOMCostRollup", _
        "SELECT" & vbCrLf & _
        "    e.RootItemID," & vbCrLf & _
        "    Sum(e.ExtQty * Nz(i.StdCost,0)) AS TotalStdCost" & vbCrLf & _
        "FROM tmpBOMExplosion AS e" & vbCrLf & _
        "INNER JOIN tblItems AS i" & vbCrLf & _
        "ON e.ComponentItemID = i.ItemID" & vbCrLf & _
        "WHERE e.RunID = TempVars!BOMRunID" & vbCrLf & _
        "GROUP BY e.RootItemID;"

    SafeUpsertQueryDef "qryBOMPrintDataset", _
        "SELECT" & vbCrLf & _
        "    h.BOMHeaderID," & vbCrLf & _
        "    fg.ItemCode AS FGCode," & vbCrLf & _
        "    fg.ItemDescription AS FGDescription," & vbCrLf & _
        "    h.VersionLabel," & vbCrLf & _
        "    h.IsActive," & vbCrLf & _
        "    e.LevelNo," & vbCrLf & _
        "    e.SortKey," & vbCrLf & _
        "    e.LineNo," & vbCrLf & _
        "    ci.ItemCode AS ComponentCode," & vbCrLf & _
        "    ci.ItemDescription AS ComponentDescription," & vbCrLf & _
        "    e.QtyPer," & vbCrLf & _
        "    e.ScrapPct," & vbCrLf & _
        "    e.ExtQty," & vbCrLf & _
        "    u.UOMCode" & vbCrLf & _
        "FROM (((tblBOMHeader AS h" & vbCrLf & _
        "INNER JOIN tblItems AS fg ON h.FGItemID = fg.ItemID)" & vbCrLf & _
        "INNER JOIN tmpBOMExplosion AS e ON h.BOMHeaderID = e.RootBOMHeaderID)" & vbCrLf & _
        "INNER JOIN tblItems AS ci ON e.ComponentItemID = ci.ItemID)" & vbCrLf & _
        "LEFT JOIN tblUOM AS u ON e.UOMID = u.UOMID" & vbCrLf & _
        "WHERE e.RunID = TempVars!BOMRunID" & vbCrLf & _
        "ORDER BY e.SortKey;"
End Sub

Private Sub SafeUpsertQueryDef(ByVal queryName As String, ByVal sqlText As String)
    On Error GoTo EH
    mCurrentStep = U("0627 06CC 062C 0627 062F 0020 06A9 0648 0626 0631 06CC 0020 0630 062E 06CC 0631 0647 200C 0634 062F 0647 003A 0020") & queryName
    UpsertQueryDef queryName, sqlText
    Exit Sub
EH:
    If IsAlreadyExistsError(Err.Number, Err.Description) Then
        Err.Clear
        Exit Sub
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub EnsureIndex(ByVal tableName As String, ByVal indexName As String, ByVal fieldsCsv As String, ByVal isUnique As Boolean)
    On Error GoTo EH
    If IndexExists(tableName, indexName) Then Exit Sub

    Dim sql As String
    sql = "CREATE " & IIf(isUnique, "UNIQUE ", "") & "INDEX " & indexName & " ON " & tableName & " (" & fieldsCsv & ");"
    ExecDDL sql
    Exit Sub

EH:
    Dim errNo As Long
    Dim errDesc As String
    errNo = Err.Number
    errDesc = Err.Description

    ' Fallback: build index by DAO API (more stable across Access variants).
    On Error Resume Next
    CreateIndexDAO tableName, indexName, fieldsCsv, isUnique
    If Err.Number = 0 Then Exit Sub
    Err.Clear
    On Error GoTo 0

    If IndexExists(tableName, indexName) Then Exit Sub
    If IsIndexErrorIgnorable(errNo, errDesc) Then Exit Sub
    Err.Raise errNo, "EnsureIndex(" & tableName & "." & indexName & ")", errDesc
End Sub

Private Sub EnsureRelation(ByVal relationName As String, ByVal parentTable As String, ByVal childTable As String, _
                           ByVal parentField As String, ByVal childField As String, _
                           Optional ByVal cascadeDelete As Boolean = False)
    Dim db As DAO.Database
    Set db = CurrentDb()

    If RelationExists(relationName) Then
        db.Relations.Delete relationName
    End If

    Dim attrs As Long
    attrs = IIf(cascadeDelete, dbRelationDeleteCascade, 0)

    Dim rel As DAO.Relation
    Set rel = db.CreateRelation(relationName, parentTable, childTable, attrs)

    Dim relField As DAO.Field
    Set relField = rel.CreateField(parentField)
    relField.ForeignName = childField
    rel.Fields.Append relField

    db.Relations.Append rel
End Sub

Private Sub UpsertQueryDef(ByVal queryName As String, ByVal sqlText As String)
    Dim db As DAO.Database
    Set db = CurrentDb()

    If QueryExists(queryName) Then
        db.QueryDefs(queryName).SQL = sqlText
    Else
        db.CreateQueryDef queryName, sqlText
    End If
End Sub

Private Sub ExecDDL(ByVal sql As String)
    CurrentDb.Execute sql, dbFailOnError
End Sub

Private Function TableExists(ByVal tableName As String) As Boolean
    Dim tdf As DAO.TableDef
    For Each tdf In CurrentDb.TableDefs
        If StrComp(tdf.Name, tableName, vbTextCompare) = 0 Then
            TableExists = True
            Exit Function
        End If
    Next tdf
End Function

Private Function IndexExists(ByVal tableName As String, ByVal indexName As String) As Boolean
    On Error GoTo SafeExit
    If Not TableExists(tableName) Then Exit Function

    Dim tdf As DAO.TableDef
    Set tdf = CurrentDb.TableDefs(tableName)

    Dim idx As DAO.Index
    For Each idx In tdf.Indexes
        If StrComp(idx.Name, indexName, vbTextCompare) = 0 Then
            IndexExists = True
            Exit Function
        End If
    Next idx

SafeExit:
End Function

Private Sub CreateIndexDAO(ByVal tableName As String, ByVal indexName As String, ByVal fieldsCsv As String, ByVal isUnique As Boolean)
    Dim db As DAO.Database
    Set db = CurrentDb()

    Dim tdf As DAO.TableDef
    Set tdf = db.TableDefs(tableName)

    Dim idx As DAO.Index
    Set idx = tdf.CreateIndex(indexName)
    idx.Unique = isUnique

    Dim parts() As String
    parts = Split(fieldsCsv, ",")

    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim fieldName As String
        fieldName = Trim$(parts(i))
        If Len(fieldName) > 0 Then
            idx.Fields.Append idx.CreateField(fieldName)
        End If
    Next i

    tdf.Indexes.Append idx
End Sub

Private Function IsAlreadyExistsError(ByVal errNo As Long, ByVal errDesc As String) As Boolean
    Dim d As String
    d = LCase$(Nz(errDesc, ""))

    If InStr(d, "already exists") > 0 Then
        IsAlreadyExistsError = True
        Exit Function
    End If

    If InStr(d, "already has an index") > 0 Then
        IsAlreadyExistsError = True
        Exit Function
    End If

    If InStr(d, "index already exists") > 0 Then
        IsAlreadyExistsError = True
        Exit Function
    End If

    If errNo = 3012 Or errNo = 3283 Or errNo = 3284 Then
        IsAlreadyExistsError = True
        Exit Function
    End If

    If errNo = 3010 Or errNo = 3004 Then
        IsAlreadyExistsError = True
        Exit Function
    End If

    If InStr(d, "duplicate") > 0 And InStr(d, "name") > 0 Then
        IsAlreadyExistsError = True
        Exit Function
    End If

    If InStr(d, "cannot create") > 0 And InStr(d, "already") > 0 Then
        IsAlreadyExistsError = True
    End If
End Function

Private Function IsIndexErrorIgnorable(ByVal errNo As Long, ByVal errDesc As String) As Boolean
    Dim d As String
    d = LCase$(Nz(errDesc, ""))

    If IsAlreadyExistsError(errNo, errDesc) Then
        IsIndexErrorIgnorable = True
        Exit Function
    End If

    If InStr(d, "index") > 0 And InStr(d, "exists") > 0 Then
        IsIndexErrorIgnorable = True
        Exit Function
    End If

    If InStr(d, "invalid procedure call or argument") > 0 Then
        IsIndexErrorIgnorable = True
        Exit Function
    End If

    If errNo = 5 Then
        IsIndexErrorIgnorable = True
    End If
End Function

Private Function ShouldIgnoreSetupError(ByVal stepName As String, ByVal errNo As Long, ByVal errDesc As String) As Boolean
    Dim s As String
    s = LCase$(Nz(stepName, ""))

    If InStr(s, "index") > 0 Or InStr(s, U("0627 06CC 0646 062F 06A9 0633")) > 0 Then
        ShouldIgnoreSetupError = IsIndexErrorIgnorable(errNo, errDesc)
        Exit Function
    End If

    If InStr(s, "relationship") > 0 Or InStr(s, U("0631 0627 0628 0637 0647")) > 0 Then
        If IsAlreadyExistsError(errNo, errDesc) Then
            ShouldIgnoreSetupError = True
            Exit Function
        End If

        If InStr(LCase$(Nz(errDesc, "")), "relationship") > 0 And InStr(LCase$(Nz(errDesc, "")), "exists") > 0 Then
            ShouldIgnoreSetupError = True
            Exit Function
        End If
    End If

    If InStr(s, "saved quer") > 0 Or InStr(s, U("06A9 0648 0626 0631 06CC")) > 0 Then
        If IsAlreadyExistsError(errNo, errDesc) Then
            ShouldIgnoreSetupError = True
            Exit Function
        End If
    End If
End Function

Private Function RelationExists(ByVal relationName As String) As Boolean
    Dim rel As DAO.Relation
    For Each rel In CurrentDb.Relations
        If StrComp(rel.Name, relationName, vbTextCompare) = 0 Then
            RelationExists = True
            Exit Function
        End If
    Next rel
End Function

Private Function QueryExists(ByVal queryName As String) As Boolean
    Dim qdf As DAO.QueryDef
    For Each qdf In CurrentDb.QueryDefs
        If StrComp(qdf.Name, queryName, vbTextCompare) = 0 Then
            QueryExists = True
            Exit Function
        End If
    Next qdf
End Function
