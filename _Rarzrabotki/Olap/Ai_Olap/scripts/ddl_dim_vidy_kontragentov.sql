-- Dim_VidyKontragentov — Справочник.А_ВидыКонтрагентовДляБаланса (_Reference56330).
-- Классификатор контрагентов для управленческого баланса (5 предопределённых:
-- Внутригрупповые/Внутренние подразделения/Собственники/Внешние/Кредиторы).
-- Snowflake от Dim_Contracts (ДоговорыКонтрагентов.А_ВидКонтрагента = _Fld56332RRef),
-- по образцу Dim_FinAgents (2026-05-19). Ключ char(32) hex-uuid (varbinary_to_uuid).
-- + 2 новые FK-колонки в Dim_Contracts:
--   VidKontragenta_ID  -> Dim_VidyKontragentov (А_ВидКонтрагента, _Fld56332RRef)
--   NapravlenieUslug_ID -> Dim_Directions      (А_НаправлениеОказаниеУслуг, _Fld56331RRef)
-- Idempotent: DROP+CREATE для Dim, guarded ALTER для Dim_Contracts. 2026-06-12.
IF OBJECT_ID('dbo.Dim_VidyKontragentov', 'U') IS NOT NULL
    DROP TABLE dbo.Dim_VidyKontragentov;
GO
CREATE TABLE dbo.Dim_VidyKontragentov (
    VidKontragenta_ID   char(32)      NOT NULL PRIMARY KEY,
    VidKontragenta_Name nvarchar(150) NULL,
    Code                varchar(20)   NULL,
    Marked_For_Deletion bit           NULL,
    Loaded_At           datetime2     NOT NULL DEFAULT SYSDATETIME()
);
GO
IF COL_LENGTH('dbo.Dim_Contracts', 'VidKontragenta_ID') IS NULL
    ALTER TABLE dbo.Dim_Contracts ADD VidKontragenta_ID char(32) NULL;
GO
IF COL_LENGTH('dbo.Dim_Contracts', 'NapravlenieUslug_ID') IS NULL
    ALTER TABLE dbo.Dim_Contracts ADD NapravlenieUslug_ID char(32) NULL;
GO
-- 2026-06-12 v2: денормализованное имя направления услуг (решение финансиста —
-- НЕ отдельное измерение, а текстовая колонка как Dept_OkazUslug_Name)
IF COL_LENGTH('dbo.Dim_Contracts', 'NapravlenieUslug_Name') IS NULL
    ALTER TABLE dbo.Dim_Contracts ADD NapravlenieUslug_Name nvarchar(150) NULL;
GO
