CREATE TABLE tblUOM (
    UOMID AUTOINCREMENT CONSTRAINT PK_tblUOM PRIMARY KEY,
    UOMCode TEXT(10) NOT NULL,
    UOMName TEXT(50),
    IsActive YESNO NOT NULL
);

CREATE UNIQUE INDEX UX_tblUOM_UOMCode ON tblUOM (UOMCode);

CREATE TABLE tblItemType (
    ItemTypeID AUTOINCREMENT CONSTRAINT PK_tblItemType PRIMARY KEY,
    TypeCode TEXT(10) NOT NULL,
    TypeName TEXT(50)
);

CREATE UNIQUE INDEX UX_tblItemType_TypeCode ON tblItemType (TypeCode);

CREATE TABLE tblItems (
    ItemID AUTOINCREMENT CONSTRAINT PK_tblItems PRIMARY KEY,
    ItemCode TEXT(30) NOT NULL,
    ItemDescription TEXT(255) NOT NULL,
    UOMID LONG,
    ItemTypeID LONG,
    IsActive YESNO NOT NULL,
    StdCost CURRENCY,
    CreatedOn DATETIME,
    CreatedBy TEXT(50),
    ModifiedOn DATETIME,
    ModifiedBy TEXT(50)
);

CREATE UNIQUE INDEX UX_tblItems_ItemCode ON tblItems (ItemCode);
CREATE INDEX IX_tblItems_UOMID ON tblItems (UOMID);
CREATE INDEX IX_tblItems_ItemTypeID ON tblItems (ItemTypeID);

CREATE TABLE tblBOMHeader (
    BOMHeaderID AUTOINCREMENT CONSTRAINT PK_tblBOMHeader PRIMARY KEY,
    FGItemID LONG NOT NULL,
    VersionNo LONG NOT NULL,
    VersionLabel TEXT(10),
    IsActive YESNO NOT NULL,
    ActiveKey LONG,
    Notes LONGTEXT,
    CreatedOn DATETIME,
    CreatedBy TEXT(50)
);

CREATE UNIQUE INDEX UX_tblBOMHeader_FG_Version ON tblBOMHeader (FGItemID, VersionNo);
CREATE UNIQUE INDEX UX_tblBOMHeader_ActiveKey ON tblBOMHeader (ActiveKey);
CREATE INDEX IX_tblBOMHeader_FGItemID ON tblBOMHeader (FGItemID);

CREATE TABLE tblBOMLines (
    BOMLineID AUTOINCREMENT CONSTRAINT PK_tblBOMLines PRIMARY KEY,
    BOMHeaderID LONG NOT NULL,
    LineNo LONG,
    ComponentItemID LONG NOT NULL,
    QtyPer DOUBLE NOT NULL,
    UOMID LONG,
    ScrapPct DOUBLE,
    EffectiveFrom DATETIME,
    EffectiveTo DATETIME,
    Notes LONGTEXT
);

CREATE UNIQUE INDEX UX_tblBOMLines_Header_Component ON tblBOMLines (BOMHeaderID, ComponentItemID);
CREATE UNIQUE INDEX UX_tblBOMLines_Header_LineNo ON tblBOMLines (BOMHeaderID, LineNo);
CREATE INDEX IX_tblBOMLines_ComponentItemID ON tblBOMLines (ComponentItemID);

CREATE TABLE tmpBOMExplosion (
    ExplosionID AUTOINCREMENT CONSTRAINT PK_tmpBOMExplosion PRIMARY KEY,
    RunID TEXT(36) NOT NULL,
    RootBOMHeaderID LONG NOT NULL,
    RootItemID LONG NOT NULL,
    ParentItemID LONG,
    ComponentItemID LONG NOT NULL,
    LevelNo LONG NOT NULL,
    LineNo LONG,
    QtyPer DOUBLE,
    ScrapPct DOUBLE,
    QtyWithScrap DOUBLE,
    ExtQty DOUBLE,
    UOMID LONG,
    SortKey TEXT(255),
    PathCodes LONGTEXT,
    AsOfDate DATETIME,
    CreatedOn DATETIME,
    CreatedBy TEXT(50)
);

CREATE INDEX IX_tmpBOMExplosion_RunID ON tmpBOMExplosion (RunID);
CREATE INDEX IX_tmpBOMExplosion_Root ON tmpBOMExplosion (RootBOMHeaderID);

ALTER TABLE tblItems
ADD CONSTRAINT FK_tblItems_tblUOM
FOREIGN KEY (UOMID) REFERENCES tblUOM (UOMID);

ALTER TABLE tblItems
ADD CONSTRAINT FK_tblItems_tblItemType
FOREIGN KEY (ItemTypeID) REFERENCES tblItemType (ItemTypeID);

ALTER TABLE tblBOMHeader
ADD CONSTRAINT FK_tblBOMHeader_tblItems_FG
FOREIGN KEY (FGItemID) REFERENCES tblItems (ItemID);

ALTER TABLE tblBOMLines
ADD CONSTRAINT FK_tblBOMLines_tblBOMHeader
FOREIGN KEY (BOMHeaderID) REFERENCES tblBOMHeader (BOMHeaderID)
ON DELETE CASCADE;

ALTER TABLE tblBOMLines
ADD CONSTRAINT FK_tblBOMLines_tblItems_Component
FOREIGN KEY (ComponentItemID) REFERENCES tblItems (ItemID);

ALTER TABLE tblBOMLines
ADD CONSTRAINT FK_tblBOMLines_tblUOM
FOREIGN KEY (UOMID) REFERENCES tblUOM (UOMID);
