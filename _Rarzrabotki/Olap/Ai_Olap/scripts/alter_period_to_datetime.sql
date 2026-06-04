-- =========================================================================
-- Stage 2026-05-06: повернути DATE → DATETIME для Period колонок
-- Calendar.date_, Fact_PnL.Period, Fact_Cashflow.Period
-- Час обрізається у Python ETL (period_offset_fix transformer + strip_time)
-- =========================================================================
USE OlapBASERP;
GO

-- 1. Calendar.date_ DATE → DATETIME (з drop/recreate PK constraint)
DECLARE @pk SYSNAME = (SELECT name FROM sys.key_constraints
                       WHERE parent_object_id=OBJECT_ID('dbo.Calendar') AND type='PK');
IF @pk IS NOT NULL EXEC('ALTER TABLE dbo.Calendar DROP CONSTRAINT [' + @pk + ']');
ALTER TABLE dbo.Calendar ALTER COLUMN date_ DATETIME NOT NULL;
ALTER TABLE dbo.Calendar ADD CONSTRAINT PK_Calendar PRIMARY KEY CLUSTERED (date_);
GO

-- 2. Fact_PnL.Period date → DATETIME
ALTER TABLE Fact_PnL ALTER COLUMN Period DATETIME NULL;
GO

-- 3. Fact_Cashflow.Period date → DATETIME
ALTER TABLE Fact_Cashflow ALTER COLUMN Period DATETIME NULL;
GO

-- Verify
SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE COLUMN_NAME IN ('date_','Period') AND TABLE_NAME IN ('Calendar','Fact_PnL','Fact_Cashflow')
ORDER BY TABLE_NAME, COLUMN_NAME;
GO
