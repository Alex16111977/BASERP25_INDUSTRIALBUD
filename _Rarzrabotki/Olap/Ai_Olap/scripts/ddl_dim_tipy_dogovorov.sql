-- Dim_TipyDogovorov — справочник-перечисление Перечисление.ТипыДоговоров (_Enum1626).
-- Источник Dim_Contracts.TipDogovora (ДоговорыКонтрагентов.ТипДоговора после
-- enum_resolver → МЕТАИМЯ). _Enum1626 имеет лишь _IDRRef/_EnumOrder (нет
-- имён/синонимов) → сидируется из 1С метаданных seed-скриптом.
-- Idempotent: DROP+CREATE (seed = DROP/CREATE + insert).
IF OBJECT_ID('dbo.Dim_TipyDogovorov', 'U') IS NOT NULL
    DROP TABLE dbo.Dim_TipyDogovorov;
GO
CREATE TABLE dbo.Dim_TipyDogovorov (
    TipDogovora      varchar(50)   COLLATE Cyrillic_General_CI_AS NOT NULL PRIMARY KEY,
    TipDogovora_Name nvarchar(100) NOT NULL,
    EnumOrder        int           NOT NULL,
    Loaded_At        datetime2     NOT NULL DEFAULT SYSDATETIME()
);
GO
