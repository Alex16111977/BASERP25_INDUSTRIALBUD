-- Dim_ObjektyRaschetov — плоский Справочник.ОбъектыРасчетов (_Reference319).
-- Источник Fact_Balance.SettlementObj_ID (Свод_РасчетыСПартнерами).
-- Recreate с реквизитами: enum ТипРасчетов/ТипОбъектаРасчетов, имена
-- Партнёр/Подразделение, composite UUID (Контрагент/Договор/Объект) +
-- тип Объекта (ТипСсылки→ИдентификаторыОбъектовМетаданных).
-- Idempotent: DROP+CREATE (loader full_reload = TRUNCATE+INSERT).
IF OBJECT_ID('dbo.Dim_ObjektyRaschetov', 'U') IS NOT NULL
    DROP TABLE dbo.Dim_ObjektyRaschetov;
GO
CREATE TABLE dbo.Dim_ObjektyRaschetov (
    SettlementObj_ID    char(32)      NOT NULL PRIMARY KEY,
    SettlementObj_Name  nvarchar(300) NOT NULL,
    Marked_For_Deletion bit           NOT NULL DEFAULT 0,
    TipRaschetov        varchar(50)   COLLATE Cyrillic_General_CI_AS NULL,
    TipObjektaRaschetov varchar(50)   COLLATE Cyrillic_General_CI_AS NULL,
    Partner_Name        nvarchar(300) NULL,
    Department_Name     nvarchar(300) NULL,
    Counterparty_ID     char(32)      NULL,
    Contract_ID         char(32)      NULL,
    Object_ID           char(32)      NULL,
    Object_Type_Name    nvarchar(300) NULL,
    Loaded_At           datetime2     NOT NULL DEFAULT SYSDATETIME()
);
GO
