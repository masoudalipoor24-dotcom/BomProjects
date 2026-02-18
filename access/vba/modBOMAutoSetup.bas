Attribute VB_Name = "modBOMAutoSetup"
Option Compare Database
Option Explicit

Private mCurrentStep As String

Public Sub AutoSetupBOM(Optional ByVal RebuildQueries As Boolean = True)
    On Error GoTo EH

    Application.Echo False
    DoCmd.Hourglass True

    mCurrentStep = "Create tables"
    EnsureTables

    mCurrentStep = "Create indexes"
    EnsureIndexes

    mCurrentStep = "Create relationships"
    EnsureRelationships

    If RebuildQueries Then
        mCurrentStep = "Create saved queries"
        EnsureSavedQueries
    End If

    mCurrentStep = "Apply defaults and validation rules"
    ApplyBOMDefaultsAndRules

    mCurrentStep = "Seed lookup data"
    SeedLookupData

    mCurrentStep = "Done"
    DoCmd.Hourglass False
    Application.Echo True
    MsgBox "BOM Auto Setup completed successfully.", vbInformation
    Exit Sub

EH:
    On Error Resume Next
    DoCmd.Hourglass False
    Application.Echo True
    MsgBox "BOM Auto Setup failed at step: " & mCurrentStep & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical
    Err.Clear
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
    EnsureIndex "tblUOM", "UX_tblUOM_UOMCode", "UOMCode", True
    EnsureIndex "tblItemType", "UX_tblItemType_TypeCode", "TypeCode", True
    EnsureIndex "tblItems", "UX_tblItems_ItemCode", "ItemCode", True
    EnsureIndex "tblItems", "IX_tblItems_UOMID", "UOMID", False
    EnsureIndex "tblItems", "IX_tblItems_ItemTypeID", "ItemTypeID", False
    EnsureIndex "tblBOMHeader", "UX_tblBOMHeader_FG_Version", "FGItemID, VersionNo", True
    EnsureIndex "tblBOMHeader", "UX_tblBOMHeader_ActiveKey", "ActiveKey", True
    EnsureIndex "tblBOMHeader", "IX_tblBOMHeader_FGItemID", "FGItemID", False
    EnsureIndex "tblBOMLines", "UX_tblBOMLines_Header_Component", "BOMHeaderID, ComponentItemID", True
    EnsureIndex "tblBOMLines", "UX_tblBOMLines_Header_LineNo", "BOMHeaderID, LineNo", True
    EnsureIndex "tblBOMLines", "IX_tblBOMLines_ComponentItemID", "ComponentItemID", False
    EnsureIndex "tmpBOMExplosion", "IX_tmpBOMExplosion_RunID", "RunID", False
    EnsureIndex "tmpBOMExplosion", "IX_tmpBOMExplosion_Root", "RootBOMHeaderID", False
End Sub

Private Sub EnsureRelationships()
    EnsureRelation "FK_tblItems_tblUOM", "tblUOM", "tblItems", "UOMID", "UOMID", False
    EnsureRelation "FK_tblItems_tblItemType", "tblItemType", "tblItems", "ItemTypeID", "ItemTypeID", False
    EnsureRelation "FK_tblBOMHeader_tblItems_FG", "tblItems", "tblBOMHeader", "ItemID", "FGItemID", False
    EnsureRelation "FK_tblBOMLines_tblBOMHeader", "tblBOMHeader", "tblBOMLines", "BOMHeaderID", "BOMHeaderID", True
    EnsureRelation "FK_tblBOMLines_tblItems_Component", "tblItems", "tblBOMLines", "ItemID", "ComponentItemID", False
    EnsureRelation "FK_tblBOMLines_tblUOM", "tblUOM", "tblBOMLines", "UOMID", "UOMID", False
End Sub

Private Sub EnsureSavedQueries()
    UpsertQueryDef "qryBOMActiveByItem", _
        "PARAMETERS pFGItemID Long;" & vbCrLf & _
        "SELECT h.*" & vbCrLf & _
        "FROM tblBOMHeader AS h" & vbCrLf & _
        "WHERE h.FGItemID = [pFGItemID]" & vbCrLf & _
        "  AND h.IsActive = True;"

    UpsertQueryDef "qryValidationDuplicates", _
        "SELECT BOMHeaderID, ComponentItemID, Count(*) AS Cnt" & vbCrLf & _
        "FROM tblBOMLines" & vbCrLf & _
        "GROUP BY BOMHeaderID, ComponentItemID" & vbCrLf & _
        "HAVING Count(*) > 1;"

    UpsertQueryDef "qryBOMLinesWithItem", _
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

    UpsertQueryDef "qryBOMExplosion", _
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

    UpsertQueryDef "qryBOMCostRollup", _
        "SELECT" & vbCrLf & _
        "    e.RootItemID," & vbCrLf & _
        "    Sum(e.ExtQty * Nz(i.StdCost,0)) AS TotalStdCost" & vbCrLf & _
        "FROM tmpBOMExplosion AS e" & vbCrLf & _
        "INNER JOIN tblItems AS i" & vbCrLf & _
        "ON e.ComponentItemID = i.ItemID" & vbCrLf & _
        "WHERE e.RunID = TempVars!BOMRunID" & vbCrLf & _
        "GROUP BY e.RootItemID;"

    UpsertQueryDef "qryBOMPrintDataset", _
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

Private Sub EnsureIndex(ByVal tableName As String, ByVal indexName As String, ByVal fieldsCsv As String, ByVal isUnique As Boolean)
    If IndexExists(tableName, indexName) Then Exit Sub

    Dim sql As String
    sql = "CREATE " & IIf(isUnique, "UNIQUE ", "") & "INDEX " & indexName & " ON " & tableName & " (" & fieldsCsv & ");"
    ExecDDL sql
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
    If Not TableExists(tableName) Then Exit Function

    Dim idx As DAO.Index
    For Each idx In CurrentDb.TableDefs(tableName).Indexes
        If StrComp(idx.Name, indexName, vbTextCompare) = 0 Then
            IndexExists = True
            Exit Function
        End If
    Next idx
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
