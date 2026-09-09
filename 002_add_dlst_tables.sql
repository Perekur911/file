-- ============================================================
-- SQLite migration: add DLST tables
-- ============================================================

PRAGMA foreign_keys = OFF;

BEGIN TRANSACTION;

-- ------------------------------------------------------------
-- DLST_results
-- ------------------------------------------------------------

CREATE TABLE DLST_results (
    resultIdDLSP       INTEGER PRIMARY KEY AUTOINCREMENT,

    sampleCode        TEXT    NOT NULL,
    TaskId            INTEGER,
    analyseNumber     INTEGER NOT NULL,

    dateStart         TEXT,
    timeStart         TEXT,
    dateEnd           TEXT,
    timeEnd           TEXT,

    pvtCell           TEXT,
    densityAP         REAL,

    Ptransfer         REAL,
    Ttransfer         REAL,
    Vtransfer         REAL,

    Pstart            REAL,
    Tstart            REAL,
    Vstart            REAL,

    density20         REAL,
    deltaMatBalance   REAL,
    sampleMassIn      REAL,
    sampleMassOut     REAL,

    dateTimeSync      TEXT,

    CONSTRAINT UQ_DLST_results_sampleCode_analyseNumber
        UNIQUE (sampleCode, analyseNumber)
);

-- ------------------------------------------------------------
-- DLST_sourceData
-- ------------------------------------------------------------

CREATE TABLE DLST_sourceData (
    rowIdDLSP              INTEGER PRIMARY KEY AUTOINCREMENT,
    resultIdDLSP           INTEGER NOT NULL,

    step                   TEXT,
    date                   TEXT,
    operator               TEXT,

    stepPcell              REAL,
    stepTcell              REAL,
    Vcell                  REAL,
    VCellLiq               REAL,
    VafterOut              REAL,

    transferPcell          REAL,
    transferTcell          REAL,

    V1                     REAL,
    V2                     REAL,

    Upic                   TEXT,
    Ppic                   TEXT,

    m0                     REAL,
    m1                     REAL,
    m0Trap                 REAL,
    m1Trap                 REAL,

    gorEquipment           TEXT,
    Patm                   REAL,
    initPatm               REAL,
    T                      REAL,

    VHeTot                 REAL,
    VHeCyl                 REAL,
    VGTot                  REAL,
    VGCyl                  REAL,
    m2                     REAL,

    sampleCodeLPh          TEXT,
    sampleCodeGPh          TEXT,

    APdensity              REAL,
    Vtransfer              REAL,
    Vgas                   REAL,
    VgasSt                 REAL,
    mgas                   REAL,
    dGas                   REAL,
    DensLiqCell            REAL,
    VLiq                   REAL,
    mLiq                   REAL,
    volumetricCoefficient  REAL,
    gasContent             REAL,
    GF                     REAL,

    CONSTRAINT FK_DLST_sourceData_resultIdDLSP
        FOREIGN KEY (resultIdDLSP)
        REFERENCES DLST_results (resultIdDLSP)
        ON UPDATE CASCADE
        ON DELETE CASCADE,

    CONSTRAINT UQ_DLST_sourceData_sampleCodeLPh
        UNIQUE (sampleCodeLPh),

    CONSTRAINT UQ_DLST_sourceData_sampleCodeGPh
        UNIQUE (sampleCodeGPh)
);

-- ------------------------------------------------------------
-- INDEXES
-- ------------------------------------------------------------

CREATE INDEX IX_DLST_results_TaskId
    ON DLST_results (TaskId);

CREATE INDEX IX_DLST_sourceData_resultIdDLSP
    ON DLST_sourceData (resultIdDLSP);

CREATE INDEX IX_DLST_sourceData_resultIdDLSP_step
    ON DLST_sourceData (resultIdDLSP, step);

COMMIT;

PRAGMA foreign_keys = ON;
