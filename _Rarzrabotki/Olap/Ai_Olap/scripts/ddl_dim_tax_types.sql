-- Dim_TaxTypes — справочник-перечисление Перечисление.ТипыНалогов (_Enum1651).
-- Источник Fact_Balance.TaxType (Свод_ПрочиеАктивыПассивы_Прямой, статья «Налоги»,
-- ВЫРАЗИТЬ ПАП.Аналитика КАК Перечисление.ТипыНалогов).
-- Ключ = МЕТАИМЯ перечисления (varchar, как хранит Fact_Balance.TaxType после
-- enum_resolver) + представление (синоним 1С UI). 14 значений + "ПустаяСсылка"
-- (строки НЕ-«Налоги»). Имена/синонимы НЕ в SQL backend (frozen enum: _Enum1651
-- имеет лишь _IDRRef/_EnumOrder) → сидируется из 1С метаданных seed-скриптом.
-- Idempotent: DROP+CREATE (seed = DROP/CREATE + insert).
IF OBJECT_ID('dbo.Dim_TaxTypes', 'U') IS NOT NULL
    DROP TABLE dbo.Dim_TaxTypes;
GO
CREATE TABLE dbo.Dim_TaxTypes (
    TaxType       varchar(50)   COLLATE Cyrillic_General_CI_AS NOT NULL PRIMARY KEY,
    TaxType_Name  nvarchar(100) NOT NULL,
    EnumOrder     int           NOT NULL,
    Loaded_At     datetime2     NOT NULL DEFAULT SYSDATETIME()
);
GO
