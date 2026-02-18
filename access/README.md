# Access BOM Implementation

This folder contains a full Microsoft Access (ACCDB) implementation package for:

- Normalized BOM data model (3NF)
- Referential integrity and indexes
- Multi-version BOM with one active version per FG/Assembly
- Duplicate prevention at line level
- Cycle detection
- Multi-level BOM explosion to a temp table
- Print dataset for reporting

## Files

- `access/sql/01_schema.sql`: tables, indexes, and foreign keys
- `access/sql/02_queries.sql`: saved query SQL definitions
- `access/vba/modBOM.bas`: core business logic (versioning, activation, copy, cycle, explosion)
- `access/vba/modBOMSchemaRules.bas`: apply defaults/validation rules + seed lookups
- `access/vba/frmItem.code.vba`: `frmItem` form events
- `access/vba/frmBOM.code.vba`: `frmBOM` main form events
- `access/vba/subfrmBOMLines.code.vba`: `subfrmBOMLines` events
- `access/vba/rptBOM.code.vba`: `rptBOM` report event

## Setup Order

1. Create a new `.accdb`.
2. Run all statements from `access/sql/01_schema.sql` (Create Query -> SQL View -> Execute).
3. Create saved queries in Access using SQL from `access/sql/02_queries.sql`.
4. Import VBA modules/code:
   - Import `modBOM.bas` as a standard module named `modBOM`.
   - Import `modBOMSchemaRules.bas` as a standard module named `modBOMSchemaRules`.
   - Paste form/report code into code-behind for matching object names.
5. In Immediate Window, run:
   - `ApplyBOMDefaultsAndRules`
   - `SeedLookupData`
6. Build forms and report using control names from the sections below.
7. Test with the scenarios in "Validation Scenarios".

## Saved Query Names

Create each query using the SQL in these files:

- `access/sql/queries/qryBOMActiveByItem.sql`
- `access/sql/queries/qryValidationDuplicates.sql`
- `access/sql/queries/qryBOMLinesWithItem.sql`
- `access/sql/queries/qryBOMExplosion.sql`
- `access/sql/queries/qryBOMCostRollup.sql`
- `access/sql/queries/qryBOMPrintDataset.sql`

## Required Form Objects and Control Names

### `frmItem`

- Record Source: `tblItems`
- Controls:
  - `txtItemCode` -> `ItemCode`
  - `txtItemDescription` -> `ItemDescription`
  - `cboUOMID` -> `UOMID`
  - `cboItemTypeID` -> `ItemTypeID`
  - `chkIsActive` -> `IsActive`
  - `txtStdCost` -> `StdCost`
  - `txtCreatedOn` -> `CreatedOn` (Locked)
  - `txtCreatedBy` -> `CreatedBy` (Locked)
  - `txtModifiedOn` -> `ModifiedOn` (Locked)
  - `txtModifiedBy` -> `ModifiedBy` (Locked)

`cboUOMID` Row Source:

```sql
SELECT UOMID, UOMCode, UOMName
FROM tblUOM
WHERE IsActive=True
ORDER BY UOMCode;
```

`cboItemTypeID` Row Source:

```sql
SELECT ItemTypeID, TypeCode, TypeName
FROM tblItemType
ORDER BY TypeCode;
```

### `frmBOM`

- Record Source: `tblBOMHeader`
- Controls:
  - `cboFGItemID` -> `FGItemID`
  - `txtVersionNo` -> `VersionNo` (Locked)
  - `txtVersionLabel` -> `VersionLabel` (Locked)
  - `chkIsActive` -> `IsActive` (Locked)
  - `cmdSave`
  - `cmdActivateBOM`
  - `cmdCopyBOM`
  - `cmdPrintBOM`
  - Subform control `sfmBOMLines` -> Source Object `subfrmBOMLines`
- Subform links:
  - `LinkMasterFields`: `BOMHeaderID`
  - `LinkChildFields`: `BOMHeaderID`

`cboFGItemID` Row Source:

```sql
SELECT i.ItemID, i.ItemCode, i.ItemDescription
FROM tblItems AS i
INNER JOIN tblItemType AS t ON i.ItemTypeID=t.ItemTypeID
WHERE i.IsActive=True
  AND t.TypeCode IN ('FG','SA')
ORDER BY i.ItemCode;
```

### `subfrmBOMLines`

- Record Source: `qryBOMLinesWithItem`
- Default View: Datasheet or Continuous Forms
- Controls:
  - `txtLineNo` -> `LineNo`
  - `cboComponentItemID` -> `ComponentItemID`
  - `txtComponentCode` -> `ComponentCode` (Locked)
  - `txtComponentDescription` -> `ComponentDescription` (Locked)
  - `txtQtyPer` -> `QtyPer`
  - `cboUOMID` -> `UOMID`
  - `txtScrapPct` -> `ScrapPct`
  - `dtEffectiveFrom` -> `EffectiveFrom`
  - `dtEffectiveTo` -> `EffectiveTo`

`cboComponentItemID` Row Source:

```sql
SELECT i.ItemID, i.ItemCode, i.ItemDescription, i.UOMID
FROM tblItems AS i
WHERE i.IsActive=True
ORDER BY i.ItemCode;
```

- `Column Count`: `4`
- `Bound Column`: `1`
- `Column Widths`: `0cm;3cm;6cm;0cm`

### `rptBOM`

- Record Source: `qryBOMPrintDataset`
- Group by: `BOMHeaderID`
- Header controls:
  - `txtFGCode`, `txtFGDesc`, `txtVersion`, `txtStatus`, `txtPrintDate`, `txtUser`
- Detail controls:
  - `txtLevelNo`, `txtComponentCode`, `txtComponentDescription`,
    `txtQtyPer`, `txtUOMCode`, `txtScrapPct`, `txtExtQty`

## Validation Scenarios

1. Create FG item + BOM v1 and add components.
2. Copy BOM to v2 and verify lines are duplicated.
3. Activate v1 then activate v2; verify only one active BOM remains.
4. Attempt cycle creation (A->B->C then C->A) and verify save is blocked.
5. Attempt duplicate component in same BOM and verify duplicate-key handling.
