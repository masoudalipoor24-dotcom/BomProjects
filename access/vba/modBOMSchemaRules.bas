Attribute VB_Name = "modBOMSchemaRules"
Option Compare Database
Option Explicit

Public Sub ApplyBOMDefaultsAndRules()
    SetFieldDefault "tblUOM", "IsActive", "True"
    ApplyDisplayCaptions

    SetFieldDefault "tblItems", "IsActive", "True"
    SetFieldDefault "tblItems", "CreatedOn", "=Now()"
    SetFieldDefaultSafe "tblItems", "CreatedBy", "=CurrentUser()", """"""
    SetFieldValidation "tblItems", "ItemCode", "Not Like ""* *""", U("06A9 062F 0020 06A9 0627 0644 0627 0020 0646 0628 0627 06CC 062F 0020 0641 0627 0635 0644 0647 0020 062F 0627 0634 062A 0647 0020 0628 0627 0634 062F 002E")

    SetFieldDefault "tblBOMHeader", "IsActive", "False"
    SetFieldDefault "tblBOMHeader", "CreatedOn", "=Now()"
    SetFieldDefaultSafe "tblBOMHeader", "CreatedBy", "=CurrentUser()", """"""
    SetTableValidation "tblBOMHeader", _
        "([IsActive]=False AND IsNull([ActiveKey])) OR ([IsActive]=True AND [ActiveKey]=[FGItemID])", _
        U("062F 0631 0020 0042 004F 004D 0020 0641 0639 0627 0644 060C 0020 0041 0063 0074 0069 0076 0065 004B 0065 0079 0020 0628 0627 06CC 062F 0020 0628 0631 0627 0628 0631 0020 0046 0047 0049 0074 0065 006D 0049 0044 0020 0628 0627 0634 062F 0020 0648 0020 062F 0631 0020 0042 004F 004D 0020 063A 06CC 0631 0641 0639 0627 0644 0020 0628 0627 06CC 062F 0020 004E 0075 006C 006C 0020 0628 0627 0634 062F 002E")

    SetFieldValidation "tblBOMLines", "QtyPer", ">0", U("0645 0642 062F 0627 0631 0020 0051 0074 0079 0050 0065 0072 0020 0628 0627 06CC 062F 0020 0628 0632 0631 06AF 200C 062A 0631 0020 0627 0632 0020 0635 0641 0631 0020 0628 0627 0634 062F 002E")
    SetFieldDefault "tblBOMLines", "ScrapPct", "0"
    SetFieldValidation "tblBOMLines", "ScrapPct", "Between 0 And 100", U("062F 0631 0635 062F 0020 0636 0627 06CC 0639 0627 062A 0020 0628 0627 06CC 062F 0020 0628 06CC 0646 0020 0030 0020 062A 0627 0020 0031 0030 0030 0020 0628 0627 0634 062F 002E")
    SetTableValidation "tblBOMLines", _
        "IsNull([EffectiveFrom]) OR IsNull([EffectiveTo]) OR [EffectiveTo] >= [EffectiveFrom]", _
        U("062A 0627 0631 06CC 062E 0020 067E 0627 06CC 0627 0646 0020 0627 0639 062A 0628 0627 0631 0020 0628 0627 06CC 062F 0020 0628 0632 0631 06AF 200C 062A 0631 0020 06CC 0627 0020 0645 0633 0627 0648 06CC 0020 062A 0627 0631 06CC 062E 0020 0634 0631 0648 0639 0020 0627 0639 062A 0628 0627 0631 0020 0628 0627 0634 062F 002E")
End Sub

Public Sub SeedLookupData()
    UpsertItemType "FG", U("06A9 0627 0644 0627 06CC 0020 0646 0647 0627 06CC 06CC")
    UpsertItemType "SA", U("0646 06CC 0645 200C 0633 0627 062E 062A 0647")
    UpsertItemType "RM", U("0645 0648 0627 062F 0020 0627 0648 0644 06CC 0647")
    UpsertItemType "PKG", U("0628 0633 062A 0647 200C 0628 0646 062F 06CC")

    UpsertUOM "PCS", U("0639 062F 062F")
    UpsertUOM "KG", U("06A9 06CC 0644 0648 06AF 0631 0645")
    UpsertUOM "M", U("0645 062A 0631")
End Sub

Private Sub SetFieldDefault(ByVal tableName As String, ByVal fieldName As String, ByVal defaultValue As String)
    Dim db As DAO.Database
    Set db = CurrentDb()
    db.TableDefs(tableName).Fields(fieldName).DefaultValue = defaultValue
End Sub

Private Sub SetFieldDefaultSafe(ByVal tableName As String, ByVal fieldName As String, ByVal primaryDefault As String, ByVal fallbackDefault As String)
    On Error GoTo Fallback
    SetFieldDefault tableName, fieldName, primaryDefault
    Exit Sub

Fallback:
    Err.Clear
    On Error Resume Next
    SetFieldDefault tableName, fieldName, fallbackDefault
    On Error GoTo 0
End Sub

Private Sub SetFieldValidation(ByVal tableName As String, ByVal fieldName As String, ByVal rule As String, ByVal message As String)
    Dim db As DAO.Database
    Set db = CurrentDb()

    db.TableDefs(tableName).Fields(fieldName).ValidationRule = rule
    db.TableDefs(tableName).Fields(fieldName).ValidationText = message
End Sub

Private Sub SetTableValidation(ByVal tableName As String, ByVal rule As String, ByVal message As String)
    Dim db As DAO.Database
    Set db = CurrentDb()

    db.TableDefs(tableName).ValidationRule = rule
    db.TableDefs(tableName).ValidationText = message
End Sub

Private Sub InsertIfMissing(ByVal tableName As String, ByVal criteria As String, ByVal insertSql As String)
    If Nz(DCount("*", tableName, criteria), 0) = 0 Then
        CurrentDb.Execute insertSql, dbFailOnError
    End If
End Sub

Private Sub UpsertItemType(ByVal typeCode As String, ByVal typeName As String)
    Dim sql As String
    sql = "UPDATE tblItemType SET TypeName=" & SqlQ(typeName) & " WHERE TypeCode=" & SqlQ(typeCode) & ";"
    CurrentDb.Execute sql, dbFailOnError

    InsertIfMissing "tblItemType", _
                    "TypeCode=" & SqlQ(typeCode), _
                    "INSERT INTO tblItemType (TypeCode, TypeName) VALUES (" & SqlQ(typeCode) & ", " & SqlQ(typeName) & ");"
End Sub

Private Sub UpsertUOM(ByVal uomCode As String, ByVal uomName As String)
    Dim sql As String
    sql = "UPDATE tblUOM SET UOMName=" & SqlQ(uomName) & ", IsActive=True WHERE UOMCode=" & SqlQ(uomCode) & ";"
    CurrentDb.Execute sql, dbFailOnError

    InsertIfMissing "tblUOM", _
                    "UOMCode=" & SqlQ(uomCode), _
                    "INSERT INTO tblUOM (UOMCode, UOMName, IsActive) VALUES (" & SqlQ(uomCode) & ", " & SqlQ(uomName) & ", True);"
End Sub

Private Function SqlQ(ByVal value As String) As String
    SqlQ = "'" & Replace$(Nz(value, ""), "'", "''") & "'"
End Function

Private Sub ApplyDisplayCaptions()
    ' Table captions for Persian UI while preserving English field names/keys.
    SetFieldCaption "tblUOM", "UOMID", U("0634 0646 0627 0633 0647 0020 0648 0627 062D 062F")
    SetFieldCaption "tblUOM", "UOMCode", U("06A9 062F 0020 0648 0627 062D 062F")
    SetFieldCaption "tblUOM", "UOMName", U("0646 0627 0645 0020 0648 0627 062D 062F")
    SetFieldCaption "tblUOM", "IsActive", U("0641 0639 0627 0644")

    SetFieldCaption "tblItemType", "ItemTypeID", U("0634 0646 0627 0633 0647 0020 0646 0648 0639")
    SetFieldCaption "tblItemType", "TypeCode", U("06A9 062F 0020 0646 0648 0639")
    SetFieldCaption "tblItemType", "TypeName", U("0646 0627 0645 0020 0646 0648 0639")

    SetFieldCaption "tblItems", "ItemID", U("0634 0646 0627 0633 0647 0020 06A9 0627 0644 0627")
    SetFieldCaption "tblItems", "ItemCode", U("06A9 062F 0020 06A9 0627 0644 0627")
    SetFieldCaption "tblItems", "ItemDescription", U("0634 0631 062D 0020 06A9 0627 0644 0627")
    SetFieldCaption "tblItems", "UOMID", U("0648 0627 062D 062F")
    SetFieldCaption "tblItems", "ItemTypeID", U("0646 0648 0639 0020 06A9 0627 0644 0627")
    SetFieldCaption "tblItems", "IsActive", U("0641 0639 0627 0644")
    SetFieldCaption "tblItems", "StdCost", U("0647 0632 06CC 0646 0647 0020 0627 0633 062A 0627 0646 062F 0627 0631 062F")
    SetFieldCaption "tblItems", "CreatedOn", U("062A 0627 0631 06CC 062E 0020 0627 06CC 062C 0627 062F")
    SetFieldCaption "tblItems", "CreatedBy", U("0627 06CC 062C 0627 062F 0020 06A9 0646 0646 062F 0647")
    SetFieldCaption "tblItems", "ModifiedOn", U("062A 0627 0631 06CC 062E 0020 0648 06CC 0631 0627 06CC 0634")
    SetFieldCaption "tblItems", "ModifiedBy", U("0648 06CC 0631 0627 06CC 0634 0020 06A9 0646 0646 062F 0647")

    SetFieldCaption "tblBOMHeader", "BOMHeaderID", U("0634 0646 0627 0633 0647 0020 0042 004F 004D")
    SetFieldCaption "tblBOMHeader", "FGItemID", U("06A9 0627 0644 0627 06CC 0020 0646 0647 0627 06CC 06CC")
    SetFieldCaption "tblBOMHeader", "VersionNo", U("0634 0645 0627 0631 0647 0020 0646 0633 062E 0647")
    SetFieldCaption "tblBOMHeader", "VersionLabel", U("0628 0631 0686 0633 0628 0020 0646 0633 062E 0647")
    SetFieldCaption "tblBOMHeader", "IsActive", U("0641 0639 0627 0644")
    SetFieldCaption "tblBOMHeader", "ActiveKey", U("06A9 0644 06CC 062F 0020 0641 0639 0627 0644")
    SetFieldCaption "tblBOMHeader", "Notes", U("062A 0648 0636 06CC 062D 0627 062A")
    SetFieldCaption "tblBOMHeader", "CreatedOn", U("062A 0627 0631 06CC 062E 0020 0627 06CC 062C 0627 062F")
    SetFieldCaption "tblBOMHeader", "CreatedBy", U("0627 06CC 062C 0627 062F 0020 06A9 0646 0646 062F 0647")

    SetFieldCaption "tblBOMLines", "BOMLineID", U("0634 0646 0627 0633 0647 0020 0631 062F 06CC 0641")
    SetFieldCaption "tblBOMLines", "BOMHeaderID", U("0634 0646 0627 0633 0647 0020 0042 004F 004D")
    SetFieldCaption "tblBOMLines", "LineNo", U("0631 062F 06CC 0641")
    SetFieldCaption "tblBOMLines", "ComponentItemID", U("0642 0637 0639 0647")
    SetFieldCaption "tblBOMLines", "QtyPer", U("0645 0635 0631 0641")
    SetFieldCaption "tblBOMLines", "UOMID", U("0648 0627 062D 062F")
    SetFieldCaption "tblBOMLines", "ScrapPct", U("0636 0627 06CC 0639 0627 062A 0020 0025")
    SetFieldCaption "tblBOMLines", "EffectiveFrom", U("0627 0632 0020 062A 0627 0631 06CC 062E")
    SetFieldCaption "tblBOMLines", "EffectiveTo", U("062A 0627 0020 062A 0627 0631 06CC 062E")
    SetFieldCaption "tblBOMLines", "Notes", U("062A 0648 0636 06CC 062D 0627 062A")
End Sub

Private Sub SetFieldCaption(ByVal tableName As String, ByVal fieldName As String, ByVal caption As String)
    On Error Resume Next
    Dim db As DAO.Database
    Set db = CurrentDb()
    db.TableDefs(tableName).Fields(fieldName).Properties("Caption") = caption
    If Err.Number <> 0 Then
        Err.Clear
        db.TableDefs(tableName).Fields(fieldName).Properties.Append _
            db.TableDefs(tableName).Fields(fieldName).CreateProperty("Caption", dbText, caption)
    End If
    On Error GoTo 0
End Sub
