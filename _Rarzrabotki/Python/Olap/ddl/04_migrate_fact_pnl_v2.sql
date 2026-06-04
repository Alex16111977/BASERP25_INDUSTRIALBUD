-- =====================================================================================
-- 04_migrate_fact_pnl_v2.sql
--
-- Schema migration for Fact_PnL after restructuring InformationRegister
-- А_ОтчетPL_Свод (Stage v2, 2026-05-05).
--
-- Changes (sync with new register structure):
--   RENAME 5 columns:
--     Sum_Plan_Grn        -> Sum_Excel
--     Sum_Plan_F1_Grn     -> Sum_F1_Excel
--     Sum_Plan_F2_Grn     -> Sum_F2_Excel
--     Sum_ERP_Grn         -> Sum_Fact     (NOT "Sum" — reserved word in SQL)
--     Source_Recorder_ID  -> Document_ID
--
--   DROP 7 columns (no longer in register):
--     Sum_Kazna_Grn, Currency_ID, Exchange_Rate, Sum_Original,
--     Source_Recorder_Type, Source_Recorder_Url, Source_Recorder_Presentation
--   (Source_Recorder_* fields now live in Dim_Documents — see 02b_dim_documents.sql)
--
--   ADD 4 columns (new dimensions/resources):
--     Sum_F1, Sum_F2 (Number split: ПКО/РКО vs other registrators)
--     Individual_ID  (FK on Dim_Individuals — for new ФизическоеЛицо dim)
--     Analytics      (nvarchar(200) — for new Аналитика dim, holds Контрагент/ФизЛицо name)
--
--   ADD 2 indexes:
--     IX_PnL_Document   ON Document_ID
--     IX_PnL_Individual ON Individual_ID
--
--   TRUNCATE Fact_PnL — old data is incompatible (Source enum changed,
--   Казна-rows are now invalid). Must re-run ETL after migration.
--
-- Idempotent: each step gated by metadata check so script can re-run safely.
-- =====================================================================================

USE OlapBASERP;
GO

PRINT '=== 04_migrate_fact_pnl_v2.sql — start ===';
GO

-- ---------------------------------------------------------------------------
-- 1. RENAME columns (sp_rename — atomic, indexes auto-update)
-- ---------------------------------------------------------------------------

IF COL_LENGTH('dbo.Fact_PnL', 'Sum_Plan_Grn') IS NOT NULL
BEGIN
    EXEC sp_rename 'dbo.Fact_PnL.Sum_Plan_Grn', 'Sum_Excel', 'COLUMN';
    PRINT '  Renamed: Sum_Plan_Grn -> Sum_Excel';
END
ELSE PRINT '  Skip rename Sum_Plan_Grn (already renamed)';
GO

IF COL_LENGTH('dbo.Fact_PnL', 'Sum_Plan_F1_Grn') IS NOT NULL
BEGIN
    EXEC sp_rename 'dbo.Fact_PnL.Sum_Plan_F1_Grn', 'Sum_F1_Excel', 'COLUMN';
    PRINT '  Renamed: Sum_Plan_F1_Grn -> Sum_F1_Excel';
END
ELSE PRINT '  Skip rename Sum_Plan_F1_Grn (already renamed)';
GO

IF COL_LENGTH('dbo.Fact_PnL', 'Sum_Plan_F2_Grn') IS NOT NULL
BEGIN
    EXEC sp_rename 'dbo.Fact_PnL.Sum_Plan_F2_Grn', 'Sum_F2_Excel', 'COLUMN';
    PRINT '  Renamed: Sum_Plan_F2_Grn -> Sum_F2_Excel';
END
ELSE PRINT '  Skip rename Sum_Plan_F2_Grn (already renamed)';
GO

IF COL_LENGTH('dbo.Fact_PnL', 'Sum_ERP_Grn') IS NOT NULL
BEGIN
    EXEC sp_rename 'dbo.Fact_PnL.Sum_ERP_Grn', 'Sum_Fact', 'COLUMN';
    PRINT '  Renamed: Sum_ERP_Grn -> Sum_Fact';
END
ELSE PRINT '  Skip rename Sum_ERP_Grn (already renamed)';
GO

IF COL_LENGTH('dbo.Fact_PnL', 'Source_Recorder_ID') IS NOT NULL
BEGIN
    EXEC sp_rename 'dbo.Fact_PnL.Source_Recorder_ID', 'Document_ID', 'COLUMN';
    PRINT '  Renamed: Source_Recorder_ID -> Document_ID';
END
ELSE PRINT '  Skip rename Source_Recorder_ID (already renamed)';
GO

-- ---------------------------------------------------------------------------
-- 2. DROP columns (no longer in register)
-- ---------------------------------------------------------------------------

IF COL_LENGTH('dbo.Fact_PnL', 'Sum_Kazna_Grn') IS NOT NULL
BEGIN
    ALTER TABLE Fact_PnL DROP COLUMN Sum_Kazna_Grn;
    PRINT '  Dropped: Sum_Kazna_Grn';
END
ELSE PRINT '  Skip drop Sum_Kazna_Grn (already dropped)';
GO

IF COL_LENGTH('dbo.Fact_PnL', 'Currency_ID') IS NOT NULL
BEGIN
    ALTER TABLE Fact_PnL DROP COLUMN Currency_ID;
    PRINT '  Dropped: Currency_ID';
END
ELSE PRINT '  Skip drop Currency_ID (already dropped)';
GO

IF COL_LENGTH('dbo.Fact_PnL', 'Exchange_Rate') IS NOT NULL
BEGIN
    ALTER TABLE Fact_PnL DROP COLUMN Exchange_Rate;
    PRINT '  Dropped: Exchange_Rate';
END
ELSE PRINT '  Skip drop Exchange_Rate (already dropped)';
GO

IF COL_LENGTH('dbo.Fact_PnL', 'Sum_Original') IS NOT NULL
BEGIN
    ALTER TABLE Fact_PnL DROP COLUMN Sum_Original;
    PRINT '  Dropped: Sum_Original';
END
ELSE PRINT '  Skip drop Sum_Original (already dropped)';
GO

IF COL_LENGTH('dbo.Fact_PnL', 'Source_Recorder_Type') IS NOT NULL
BEGIN
    ALTER TABLE Fact_PnL DROP COLUMN Source_Recorder_Type;
    PRINT '  Dropped: Source_Recorder_Type';
END
ELSE PRINT '  Skip drop Source_Recorder_Type (already dropped)';
GO

IF COL_LENGTH('dbo.Fact_PnL', 'Source_Recorder_Url') IS NOT NULL
BEGIN
    ALTER TABLE Fact_PnL DROP COLUMN Source_Recorder_Url;
    PRINT '  Dropped: Source_Recorder_Url';
END
ELSE PRINT '  Skip drop Source_Recorder_Url (already dropped)';
GO

IF COL_LENGTH('dbo.Fact_PnL', 'Source_Recorder_Presentation') IS NOT NULL
BEGIN
    ALTER TABLE Fact_PnL DROP COLUMN Source_Recorder_Presentation;
    PRINT '  Dropped: Source_Recorder_Presentation';
END
ELSE PRINT '  Skip drop Source_Recorder_Presentation (already dropped)';
GO

-- ---------------------------------------------------------------------------
-- 3. ADD new columns
-- ---------------------------------------------------------------------------

IF COL_LENGTH('dbo.Fact_PnL', 'Sum_F1') IS NULL
BEGIN
    ALTER TABLE Fact_PnL ADD Sum_F1 decimal(15,2) NOT NULL DEFAULT 0;
    PRINT '  Added: Sum_F1 decimal(15,2) DEFAULT 0';
END
ELSE PRINT '  Skip add Sum_F1 (already exists)';
GO

IF COL_LENGTH('dbo.Fact_PnL', 'Sum_F2') IS NULL
BEGIN
    ALTER TABLE Fact_PnL ADD Sum_F2 decimal(15,2) NOT NULL DEFAULT 0;
    PRINT '  Added: Sum_F2 decimal(15,2) DEFAULT 0';
END
ELSE PRINT '  Skip add Sum_F2 (already exists)';
GO

IF COL_LENGTH('dbo.Fact_PnL', 'Individual_ID') IS NULL
BEGIN
    ALTER TABLE Fact_PnL ADD Individual_ID char(32) NULL;
    PRINT '  Added: Individual_ID char(32) NULL';
END
ELSE PRINT '  Skip add Individual_ID (already exists)';
GO

IF COL_LENGTH('dbo.Fact_PnL', 'Analytics') IS NULL
BEGIN
    ALTER TABLE Fact_PnL ADD Analytics nvarchar(200) NULL;
    PRINT '  Added: Analytics nvarchar(200) NULL';
END
ELSE PRINT '  Skip add Analytics (already exists)';
GO

-- ---------------------------------------------------------------------------
-- 4. ADD indexes on new FK columns
-- ---------------------------------------------------------------------------

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_PnL_Document' AND object_id = OBJECT_ID('Fact_PnL'))
BEGIN
    CREATE INDEX IX_PnL_Document ON Fact_PnL (Document_ID);
    PRINT '  Created: IX_PnL_Document ON Fact_PnL (Document_ID)';
END
ELSE PRINT '  Skip create IX_PnL_Document (already exists)';
GO

IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'IX_PnL_Individual' AND object_id = OBJECT_ID('Fact_PnL'))
BEGIN
    CREATE INDEX IX_PnL_Individual ON Fact_PnL (Individual_ID);
    PRINT '  Created: IX_PnL_Individual ON Fact_PnL (Individual_ID)';
END
ELSE PRINT '  Skip create IX_PnL_Individual (already exists)';
GO

-- ---------------------------------------------------------------------------
-- 5. TRUNCATE — old data is incompatible (Source enum changed, Казна rows invalid)
-- ---------------------------------------------------------------------------

PRINT '  Pre-truncate row count:';
SELECT COUNT(*) AS rows_before_truncate FROM Fact_PnL;

TRUNCATE TABLE Fact_PnL;

PRINT '  Post-truncate row count:';
SELECT COUNT(*) AS rows_after_truncate FROM Fact_PnL;
GO

-- ---------------------------------------------------------------------------
-- 6. Verification
-- ---------------------------------------------------------------------------

PRINT '=== Final Fact_PnL schema ===';
SELECT
    ORDINAL_POSITION AS pos,
    COLUMN_NAME      AS col,
    DATA_TYPE        AS type,
    CHARACTER_MAXIMUM_LENGTH AS len,
    IS_NULLABLE      AS nullable
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'Fact_PnL'
ORDER BY ORDINAL_POSITION;

PRINT '=== Indexes on Fact_PnL ===';
SELECT name FROM sys.indexes WHERE object_id = OBJECT_ID('Fact_PnL') AND name IS NOT NULL ORDER BY name;

PRINT '=== 04_migrate_fact_pnl_v2.sql — done ===';
GO
