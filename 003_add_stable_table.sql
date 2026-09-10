-- ============================================================
-- SQLite migration: add Stable table
-- ============================================================

PRAGMA foreign_keys = OFF;

BEGIN TRANSACTION;

-- ------------------------------------------------------------
-- Stable_results
-- ------------------------------------------------------------

CREATE TABLE Stable_results (
    resultIdStable INTEGER PRIMARY KEY AUTOINCREMENT,

    TaskId         INTEGER,
    sampleCode     TEXT    NOT NULL,
    analyseNumber  INTEGER NOT NULL,

    dateStart      TEXT,
    timeStart      TEXT,
    dateEnd        TEXT,
    timeEnd        TEXT,

    operator       TEXT,
    PVTcell        TEXT,

    temperature    REAL,
    pressure       REAL,
    pressureUnit   TEXT,
    pressureAbs    INTEGER,
    pressureMPaAbs REAL,
    LiquidVolume   REAL,

    dateTimeSync   TEXT
);

-- ------------------------------------------------------------
-- INDEXES
-- ------------------------------------------------------------

-- Для одной пробы допускаются разные номера анализа,
-- но одна и та же пара sampleCode + analyseNumber уникальна.
CREATE UNIQUE INDEX UX_Stable_results_sampleCode_analyseNumber
    ON Stable_results (sampleCode, analyseNumber);

-- Используется при поиске результата по заданию.
CREATE INDEX IX_Stable_results_TaskId
    ON Stable_results (TaskId);

COMMIT;

PRAGMA foreign_keys = ON;
