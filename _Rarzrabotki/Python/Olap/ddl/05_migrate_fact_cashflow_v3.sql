-- =====================================================================================
-- 05_migrate_fact_cashflow_v3.sql
--
-- Schema migration for Fact_Cashflow after restructuring InformationRegister
-- А_ОтчетDDS_Свод (Stage v3, 2026-05-08).
--
-- New register structure: 14 dim + 5 res + 1 attr.
--   Dimensions: ВидДвижения, ТипДенежныхСредств, Источник, ХозяйственнаяОперация,
--               Организация, Подразделение, ДенежныеСредства (composite),
--               СтатьяДвиженияДенежныхСредств, Контрагент, Договор, ФизическоеЛицо,
--               Валюта, ДокументДвижения (composite), Аналитика (Строка).
--   Resources: Сумма, СуммаВВалюте, СуммаКазна, СуммаПлан, СуммаПланОбъект.
--   Attribute: Комментарий.
--   Source enum: 4 values (ЕРП / Казна / План / ПланОбъекта).
--
-- Strategy: DROP + CREATE (no data preserved — old structure incompatible).
--
-- Idempotent: each step gated by existence check.
-- =====================================================================================

USE OlapBASERP;
GO

PRINT '=== 05_migrate_fact_cashflow_v3.sql — start ===';
GO

-- ---------------------------------------------------------------------------
-- 1. DROP existing Fact_Cashflow (Stage 1 structure: 18 cols + 4 drill-down)
-- ---------------------------------------------------------------------------

IF OBJECT_ID('dbo.Fact_Cashflow', 'U') IS NOT NULL
BEGIN
    PRINT '  Pre-drop row count:';
    SELECT COUNT(*) AS rows_before_drop FROM Fact_Cashflow;

    DROP TABLE dbo.Fact_Cashflow;
    PRINT '  Dropped: Fact_Cashflow (Stage 1)';
END
ELSE
    PRINT '  Skip drop Fact_Cashflow (does not exist)';
GO

-- ---------------------------------------------------------------------------
-- 2. CREATE new Fact_Cashflow (Stage v3: 14 dim + 5 res + Comment + Analytics)
-- ---------------------------------------------------------------------------

CREATE TABLE dbo.Fact_Cashflow (
    Fact_ID                     bigint IDENTITY(1,1) PRIMARY KEY,

    -- Period + Recorder
    Period                      datetime2 NOT NULL,
    Period_Month                date NOT NULL,
    Recorder_DDS_ID             char(32) NULL,        -- UUID Документ.А_ФинРез_DDS

    -- 14 dimensions (1:1 mapping with А_ОтчетDDS_Свод)
    Movement_Type               varchar(15)  NULL,    -- enum: Поступление|Списание
    Cash_Type                   varchar(40)  NULL,    -- enum: Наличные|Безналичные|...
    Source                      varchar(20)  NOT NULL,-- enum А_ИсточникDDS: ЕРП|Казна|План|ПланОбъекта
    Business_Operation          varchar(80)  NULL,    -- enum ХозОперация (~600 values, dynamic resolver)
    Organization_ID             char(32) NOT NULL,
    Department_ID               char(32) NULL,
    Cash_Account_ID             char(32) NULL,        -- FK Dim_DenezhnyeSredstva (composite)
    DDS_Article_ID              char(32) NULL,
    Counterparty_ID             char(32) NULL,
    Contract_ID                 char(32) NULL,
    Individual_ID               char(32) NULL,
    Currency_ID                 char(32) NULL,
    Document_ID                 char(32) NULL,        -- FK Dim_Documents (shared with Fact_PnL)
    Analytics                   nvarchar(150) NULL,   -- Строка (denormalized)

    -- Attribute
    Comment                     nvarchar(1000) NULL,

    -- 5 resources
    Sum_Fact                    decimal(15,2) NOT NULL DEFAULT 0,  -- 1С: Сумма
    Sum_Currency                decimal(15,2) NOT NULL DEFAULT 0,  -- 1С: СуммаВВалюте
    Sum_Kazna                   decimal(15,2) NOT NULL DEFAULT 0,  -- 1С: СуммаКазна
    Sum_Plan                    decimal(15,2) NOT NULL DEFAULT 0,  -- 1С: СуммаПлан
    Sum_Plan_Object             decimal(15,2) NOT NULL DEFAULT 0,  -- 1С: СуммаПланОбъект

    Loaded_At                   datetime2 NOT NULL DEFAULT SYSDATETIME(),

    INDEX IX_CF_Period_Source  (Period_Month, Source),
    INDEX IX_CF_Department     (Department_ID, Period_Month),
    INDEX IX_CF_DDS_Article    (DDS_Article_ID, Period_Month),
    INDEX IX_CF_Counterparty   (Counterparty_ID, Period_Month),
    INDEX IX_CF_Document       (Document_ID),
    INDEX IX_CF_Cash_Account   (Cash_Account_ID, Period_Month)
);

PRINT '  Created: Fact_Cashflow (Stage v3, 21 columns)';
GO

-- ---------------------------------------------------------------------------
-- 3. Verification
-- ---------------------------------------------------------------------------

PRINT '=== Final Fact_Cashflow schema ===';
SELECT
    ORDINAL_POSITION AS pos,
    COLUMN_NAME      AS col,
    DATA_TYPE        AS type,
    CHARACTER_MAXIMUM_LENGTH AS len,
    IS_NULLABLE      AS nullable
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'Fact_Cashflow'
ORDER BY ORDINAL_POSITION;

PRINT '=== Indexes on Fact_Cashflow ===';
SELECT name FROM sys.indexes WHERE object_id = OBJECT_ID('Fact_Cashflow') AND name IS NOT NULL ORDER BY name;

PRINT '=== 05_migrate_fact_cashflow_v3.sql — done ===';
GO
