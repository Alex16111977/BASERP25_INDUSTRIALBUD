-- Dim_TipPokazatelya — справочник-перечисление
-- Перечисление.ВидыСтатейУправленческогоБаланса (Актив/Пассив/АктивПассив).
-- Источник Fact_Balance.TipPokazatelya = измерение А_ОтчетБаланс_Свод.
-- ТипПоказателя, заполняемое в ПровестиБалансСвод по формуле штатного
-- Отчет.УправленческийБаланс (АктивПассив→Пассив; в регистр пишется
-- только Актив/Пассив). Ключ = МЕТАИМЯ enum (varchar, как хранит
-- Fact_Balance.TipPokazatelya после enum_resolver) + представление 1С UI.
-- frozen enum (_EnumOrder), имена/синонимы сидируются из 1С метаданных.
-- Idempotent: DROP+CREATE (seed = DROP/CREATE + insert) + ALTER Fact_Balance.
IF COL_LENGTH('dbo.Fact_Balance', 'TipPokazatelya') IS NULL
    ALTER TABLE dbo.Fact_Balance
        ADD TipPokazatelya varchar(50) COLLATE Cyrillic_General_CI_AS NULL;
GO
IF OBJECT_ID('dbo.Dim_TipPokazatelya', 'U') IS NOT NULL
    DROP TABLE dbo.Dim_TipPokazatelya;
GO
CREATE TABLE dbo.Dim_TipPokazatelya (
    TipPokazatelya      varchar(50)   COLLATE Cyrillic_General_CI_AS NOT NULL PRIMARY KEY,
    TipPokazatelya_Name nvarchar(100) NOT NULL,
    EnumOrder           int           NOT NULL,
    Loaded_At           datetime2     NOT NULL DEFAULT SYSDATETIME()
);
GO
