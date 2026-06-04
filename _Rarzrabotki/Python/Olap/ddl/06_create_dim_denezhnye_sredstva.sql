-- =====================================================================================
-- 06_create_dim_denezhnye_sredstva.sql
--
-- Dim_DenezhnyeSredstva — universal composite dimension for Fact_Cashflow.Cash_Account_ID.
-- Source: UNION of 8 1С catalogs (Cash_Type=composite in regis А_ОтчетDDS_Свод):
--   1. БанковскиеСчетаОрганизаций (БанковскийСчет)
--   2. Кассы                       (Касса)
--   3. КассыККМ                    (КассаККМ)
--   4. ДенежныеДокументы           (ДенежныйДокумент)
--   5. ФизическиеЛица              (ФизическоеЛицо)
--   6. ДоговорыКредитовИДепозитов  (ДоговорКредитаДепозита)
--   7. Контрагенты                 (Контрагент)
--   8. ЭквайринговыеТерминалы      (ЭквайринговыйТерминал)
-- (rename 2026-05-20: было BankAccount/Cashbox/...; varchar(30) Cyrillic_General_CI_AS)
--
-- ETL via pipelines/dim_denezhnye_sredstva.json (UNION 8 SELECTs from BaseERP _Reference*).
--
-- Idempotent: DROP + CREATE.
-- =====================================================================================

USE OlapBASERP;
GO

PRINT '=== 06_create_dim_denezhnye_sredstva.sql — start ===';
GO

IF OBJECT_ID('dbo.Dim_DenezhnyeSredstva', 'U') IS NOT NULL
BEGIN
    DROP TABLE dbo.Dim_DenezhnyeSredstva;
    PRINT '  Dropped existing Dim_DenezhnyeSredstva';
END
GO

CREATE TABLE dbo.Dim_DenezhnyeSredstva (
    Cash_Account_ID  char(32) PRIMARY KEY,
    Account_Name     nvarchar(250) NOT NULL,
    Account_Type     varchar(30)   NOT NULL,
    Organization_ID  char(32)      NULL,
    Currency_Code    char(3)       NULL,
    Direction_ID     char(32)      NULL,  -- А_НаправлениеДеятельности (Касса._Fld54708RRef / БанковскийСчет._Fld54618RRef); 2026-05-20
    AccountingOrg_ID char(32)      NULL,  -- А_ОрганизацияБухгалтерия (Касса._Fld55471RRef / БанковскийСчет._Fld54501RRef); 2026-05-20. FK→Dim_Organizations (полный, БЕЗ фильтра own-org)
    Loaded_At        datetime2     NOT NULL DEFAULT SYSDATETIME(),
    INDEX IX_CashAcc_Type (Account_Type)
);
PRINT '  Created: Dim_DenezhnyeSredstva';
GO

-- Unknown member (for orphan FK fallback)
INSERT INTO dbo.Dim_DenezhnyeSredstva (Cash_Account_ID, Account_Name, Account_Type, Organization_ID, Currency_Code)
VALUES (REPLICATE('0', 32), N'(не вказано)', N'НеИзвестно', NULL, NULL);
PRINT '  Inserted unknown member';
GO

PRINT '=== Final Dim_DenezhnyeSredstva schema ===';
SELECT
    ORDINAL_POSITION AS pos,
    COLUMN_NAME      AS col,
    DATA_TYPE        AS type,
    CHARACTER_MAXIMUM_LENGTH AS len,
    IS_NULLABLE      AS nullable
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'Dim_DenezhnyeSredstva'
ORDER BY ORDINAL_POSITION;

PRINT '=== 06_create_dim_denezhnye_sredstva.sql — done ===';
GO
