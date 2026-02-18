Attribute VB_Name = "modBOMSchemaRules"
Option Compare Database
Option Explicit

Public Sub ApplyBOMDefaultsAndRules()
    SetFieldDefault "tblUOM", "IsActive", "True"

    SetFieldDefault "tblItems", "IsActive", "True"
    SetFieldDefault "tblItems", "CreatedOn", "=Now()"
    SetFieldDefaultSafe "tblItems", "CreatedBy", "=CurrentUser()", """"""
    SetFieldValidation "tblItems", "ItemCode", "Not Like ""* *""", "ItemCode cannot contain spaces."

    SetFieldDefault "tblBOMHeader", "IsActive", "False"
    SetFieldDefault "tblBOMHeader", "CreatedOn", "=Now()"
    SetFieldDefaultSafe "tblBOMHeader", "CreatedBy", "=CurrentUser()", """"""
    SetTableValidation "tblBOMHeader", _
        "([IsActive]=False AND IsNull([ActiveKey])) OR ([IsActive]=True AND [ActiveKey]=[FGItemID])", _
        "Active BOM must have ActiveKey = FGItemID; inactive BOM must have ActiveKey = Null."

    SetFieldValidation "tblBOMLines", "QtyPer", ">0", "QtyPer must be greater than zero."
    SetFieldDefault "tblBOMLines", "ScrapPct", "0"
    SetFieldValidation "tblBOMLines", "ScrapPct", "Between 0 And 100", "ScrapPct must be between 0 and 100."
    SetTableValidation "tblBOMLines", _
        "IsNull([EffectiveFrom]) OR IsNull([EffectiveTo]) OR [EffectiveTo] >= [EffectiveFrom]", _
        "EffectiveTo must be greater than or equal to EffectiveFrom."
End Sub

Public Sub SeedLookupData()
    InsertIfMissing "tblItemType", "TypeCode='FG'", "INSERT INTO tblItemType (TypeCode, TypeName) VALUES ('FG','Finished Good')"
    InsertIfMissing "tblItemType", "TypeCode='SA'", "INSERT INTO tblItemType (TypeCode, TypeName) VALUES ('SA','Sub Assembly')"
    InsertIfMissing "tblItemType", "TypeCode='RM'", "INSERT INTO tblItemType (TypeCode, TypeName) VALUES ('RM','Raw Material')"
    InsertIfMissing "tblItemType", "TypeCode='PKG'", "INSERT INTO tblItemType (TypeCode, TypeName) VALUES ('PKG','Packaging')"

    InsertIfMissing "tblUOM", "UOMCode='PCS'", "INSERT INTO tblUOM (UOMCode, UOMName, IsActive) VALUES ('PCS','Piece',True)"
    InsertIfMissing "tblUOM", "UOMCode='KG'", "INSERT INTO tblUOM (UOMCode, UOMName, IsActive) VALUES ('KG','Kilogram',True)"
    InsertIfMissing "tblUOM", "UOMCode='M'", "INSERT INTO tblUOM (UOMCode, UOMName, IsActive) VALUES ('M','Meter',True)"
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
