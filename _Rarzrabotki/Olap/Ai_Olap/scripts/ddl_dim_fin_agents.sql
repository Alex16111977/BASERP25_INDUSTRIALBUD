-- Dim_FinAgents — плоский Справочник.А_ФинАгенты (_Reference54722).
-- Источник Dim_Contracts.FinAgent_ID (ДоговорыКонтрагентов.А_ФинАгент).
-- Каждый 1С справочник имеет _IDRRef/_Description/_Marked → плоский.
-- +unknown-member 0x..01 '(Пусто)' (FK-цель). Idempotent: DROP+CREATE
-- (loader full_reload = TRUNCATE+INSERT).
IF OBJECT_ID('dbo.Dim_FinAgents', 'U') IS NOT NULL
    DROP TABLE dbo.Dim_FinAgents;
GO
CREATE TABLE dbo.Dim_FinAgents (
    FinAgent_ID         char(32)      NOT NULL PRIMARY KEY,
    FinAgent_Name       nvarchar(300) NOT NULL,
    Marked_For_Deletion bit           NOT NULL DEFAULT 0,
    Loaded_At           datetime2     NOT NULL DEFAULT SYSDATETIME()
);
GO
