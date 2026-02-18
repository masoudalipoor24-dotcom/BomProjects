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
- `access/vba/modBOMAutoSetup.bas`: one-click installer (tables, indexes, relations, queries, rules, seed)
- `access/vba/modBOMUIBuilder.bas`: one-click form/report builder inside Access
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
   - Import `modBOMAutoSetup.bas` as a standard module named `modBOMAutoSetup`.
   - Import `modBOMUIBuilder.bas` as a standard module named `modBOMUIBuilder`.
   - Paste form/report code into code-behind for matching object names.
5. In Immediate Window, run one command:
   - `AutoSetupBOM`
6. Build forms and report using control names from the sections below.
7. Test with the scenarios in "Validation Scenarios".

## One-Click Auto Setup

`AutoSetupBOM` does all backend setup in one run:

- Creates missing tables
- Creates required indexes (including unique constraints)
- Creates/refreshes relationships (with cascade delete only for `tblBOMHeader -> tblBOMLines`)
- Creates/updates saved queries
- Applies validation/default rules
- Seeds baseline lookup data (`tblItemType`, `tblUOM`)

Run options:

1. Immediate Window: `AutoSetupBOM`
2. Macro `RunCode`: `RunAutoSetupBOM()`
3. Button click event: `Call AutoSetupBOM`

## One-Click UI Build (Forms + Report)

After importing modules, run:

```vb
AutoSetupBOMWithUI
```

This will:

- run backend setup (`AutoSetupBOM`)
- create/replace `frmItem`
- create/replace `subfrmBOMLines`
- create/replace `frmBOM`
- create/replace `rptBOM`
- attach code from files in `access/vba/*.code.vba` to created objects

Alternative command (same result):

```vb
AutoSetupAllWithUI
```

## Populate Existing ACCDB from Terminal

If you already have an `.accdb` file and want to inject the BOM schema from terminal:

```powershell
python access/scripts/populate_accdb.py --db access/BomProjects.accdb --schema access/sql/01_schema.sql
```

This script creates tables/indexes/FKs from SQL and seeds lookup data (`tblUOM`, `tblItemType`).
It is idempotent on re-run:

- existing objects are counted as `SCHEMA_SKIP`
- only real unexpected errors are reported as `SCHEMA_FAIL`

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
